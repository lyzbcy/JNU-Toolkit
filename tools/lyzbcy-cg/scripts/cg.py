"""Inspect CG homework through an authenticated Edge tab."""
import argparse
import getpass
import imaplib
import ssl
from datetime import date, timedelta
from email import policy
from email.parser import BytesParser
from html.parser import HTMLParser
import os
import subprocess
from update import schedule, staged_command
import json
from pathlib import Path
import re
import sys
import time
from urllib.error import URLError
from urllib.parse import urljoin, urlsplit
from urllib.request import Request, urlopen

HERE = Path(__file__).resolve().parent
COURSE_LINK = '.courselist a[href*="courseID="]'


class BridgeError(Exception):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


class Bridge:
    def __init__(self, address):
        self.address = address.rstrip("/") + "/command"

    def call(self, action, **args):
        payload = json.dumps({"action": action, "args": args, "session": "cg"}).encode("utf-8")
        request = Request(self.address, payload, {"Content-Type": "application/json"})
        with urlopen(request, timeout=45) as response:
            reply = json.load(response)
        if not reply.get("ok"):
            error = reply["error"]
            raise BridgeError(error["code"], error["message"])
        return reply["data"]

    def cdp(self, method, **params):
        return self.call("cdp", method=method, params=params)

    def evaluate(self, expression, context=None):
        params = {"expression": expression, "returnByValue": True}
        if context is not None:
            params["contextId"] = context
        result = self.cdp("Runtime.evaluate", **params)
        if "exceptionDetails" in result:
            raise BridgeError("page_error", result["exceptionDetails"]["text"])
        return result["result"].get("value")


def default_url():
    configured = os.getenv("CG_WORKSPACE_URL")
    if configured:
        return configured
    for source in (Path.cwd()/"agents.md", HERE.parents[1]/"agents.md"):
        if source.is_file():
            notes = source.read_text(encoding="utf-8-sig")
            match = re.search(r'https://[^\s]+/courselist\.jsp\?courseID=0', notes)
            if match:
                return match.group()
    raise BridgeError("missing_url", "请通过 --url 提供 CG 课程列表地址。")


def open_courses(bridge, url, base):
    try:
        tab = bridge.call("find_tab", url=url)
    except BridgeError as error:
        if "no tab matching" not in str(error).lower():
            raise
        tab = None
    if not tab or tab.get("url") != url:
        navigate(bridge, url, not tab)

    read_page = (
        "({url:location.href,count:document.querySelectorAll("
        + json.dumps(COURSE_LINK) + ").length})"
    )
    state = bridge.evaluate(read_page)
    if not state["count"] and "/authserver" not in state["url"]:
        # The platform home establishes its own login after WebVPN authentication.
        navigate(bridge, base, False)
        state = bridge.evaluate(read_page)
    if state["count"]:
        return state
    if "/authserver" in state["url"] or "/indexcs/simple.jsp" in state["url"]:
        raise BridgeError("not_logged_in", "请在 Edge 完成 CG 登录，看到「选择课程」后重试。")
    raise BridgeError("course_list_unavailable", "CG 课程列表未加载，请检查 Edge 中的 CG 页面。")


def navigate(bridge, url, new_tab):
    try:
        bridge.call("navigate", url=url, newTab=new_tab, group_title="CG 作业")
    except BridgeError as error:
        if "page load timeout" not in str(error):
            raise
        # A slow asset can keep navigation pending after the document is already usable.
        tabs = bridge.call("list_tabs")["tabs"]
        matches = [tab for tab in tabs if urlsplit(tab["url"]).netloc == urlsplit(url).netloc]
        if not matches:
            raise
        bridge.call("attach_tab", tabId=matches[-1]["tabId"])


def scan(bridge, base):
    frame = bridge.cdp("Page.getFrameTree")["frameTree"]["frame"]
    context = bridge.cdp("Page.createIsolatedWorld", frameId=frame["id"],
                         worldName="cg-homework")["executionContextId"]
    reader = (HERE / "scan.js").read_text(encoding="utf-8")
    bridge.evaluate("void (" + reader + ")(" + json.dumps(base) + ")", context)
    deadline = time.monotonic() + 150
    while time.monotonic() < deadline:
        state = bridge.evaluate("globalThis.cgHomeworkScan ?? null", context)
        if state and state["done"]:
            return state
        time.sleep(2)
    raise BridgeError("scan_timeout", "课程扫描超时。请检查 Edge 中的 CG 页面后重试。")


MAIL_HOST = "imap.exmail.qq.com"  # Official campus guide supports this common endpoint.
MAIL_DOMAINS = {"stu.jiangnan.edu.cn", "jiangnan.edu.cn", "vip.jiangnan.edu.cn"}


class MailError(Exception):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


def valid_mail_account(account):
    name, sep, domain = account.strip().rpartition("@")
    return bool(sep and name and not any(c.isspace() for c in name)
                and domain.lower() in MAIL_DOMAINS)


def mail_since(days, today=None):
    today = today or date.today()
    start = today - timedelta(days=days - 1)
    months = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    return f"{start.day:02d}-{months[start.month-1]}-{start.year}"


def literal_payload(reply):
    status, parts = reply
    if status != "OK":
        raise MailError("mail_fetch_failed", "邮件预览读取失败；请稍后重试。")
    return next((x[1] for x in parts if isinstance(x, tuple) and isinstance(x[1], bytes)), None)


def parse_header(raw):
    try:
        msg = BytesParser(policy=policy.default).parsebytes(raw)
        return {key: str(msg.get(key, "")) for key in ("Subject", "From", "Date")}
    except (ValueError, TypeError, UnicodeError):
        raise MailError("mail_parse_failed", "邮件头格式无法解析。") from None


class PlainHTML(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self.skip += 1
        elif tag in ("br", "p", "div", "li", "tr"):
            self.parts.append(" ")

    def handle_endtag(self, tag):
        if tag in ("script", "style") and self.skip:
            self.skip -= 1
        elif tag in ("p", "div", "li", "tr"):
            self.parts.append(" ")

    def handle_data(self, data):
        if not self.skip:
            self.parts.append(data)


def body_preview(raw, max_chars):
    try:
        msg = BytesParser(policy=policy.default).parsebytes(raw)
        part = msg.get_body(preferencelist=("plain", "html"))
        if part is None or part.get_content_disposition() == "attachment":
            return ""
        content = part.get_content()
        if not isinstance(content, str):
            return ""
        if part.get_content_type() == "text/html":
            parser = PlainHTML()
            parser.feed(content)
            content = "".join(parser.parts)
        return " ".join(content.split())[:max_chars]
    except (ValueError, TypeError, UnicodeError, LookupError):
        return ""


def read_mail(account, password, days=7, limit=20, scan_limit=150, keyword="",
              preview_chars=0, client_factory=imaplib.IMAP4_SSL):
    """Read the owner's inbox in EXAMINE/read-only mode; no stores or sends."""
    if not valid_mail_account(account) or not password:
        raise MailError("mail_credentials_required", "请使用本人江南大学完整邮箱地址与已授权邮箱密码/客户端专用密码。")
    if not (1 <= days <= 366 and 1 <= limit <= 200 and limit <= scan_limit <= 1000
            and 0 <= preview_chars <= 4000):
        raise MailError("mail_bad_range", "天数为1—366、结果1—200、扫描数不少于结果数且最多1000，正文预览最多4000字符。")
    client = None
    try:
        client = client_factory(MAIL_HOST, ssl_context=ssl.create_default_context(), timeout=15)
        try:
            client.login(account, password)
        except imaplib.IMAP4.error:
            raise MailError("mail_login_failed", "邮箱登录失败。请确认本人已激活邮箱并按官方说明开启 IMAP；需要时使用客户端专用密码。") from None
        status, _ = client.select("INBOX", readonly=True)
        if status != "OK":
            raise MailError("mail_inbox_unavailable", "收件箱不可读，未取得邮件。")
        status, result = client.uid("SEARCH", None, "SINCE", mail_since(days))
        if status != "OK":
            raise MailError("mail_search_failed", "邮箱日期检索失败。")
        uids = sorted({uid for row in result if isinstance(row, bytes) for uid in row.split()
                       if uid.isdigit()}, key=int, reverse=True)
        messages, scanned, skipped = [], 0, 0
        wanted = keyword.casefold()
        for uid in uids[:scan_limit]:
            scanned += 1
            raw = literal_payload(client.uid("FETCH", uid, "(BODY.PEEK[HEADER.FIELDS (DATE FROM SUBJECT)])"))
            if raw is None:
                skipped += 1
                continue
            values = parse_header(raw)
            if wanted and wanted not in (values["Subject"] + " " + values["From"]).casefold():
                continue
            item = {"uid": uid.decode("ascii"), "subject": values["Subject"],
                    "from": values["From"], "date": values["Date"]}
            if preview_chars:
                partial = literal_payload(client.uid("FETCH", uid, "(BODY.PEEK[]<0.65536>)"))
                item["preview"] = body_preview(partial, preview_chars) if partial else ""
                item["preview_is_bounded"] = True
            messages.append(item)
            if len(messages) >= limit:
                break
        return {"scope": "INBOX", "read_only": True, "days": days,
                "matching_keyword_scope": "subject/from", "available_in_date_range": len(uids),
                "scanned": scanned, "skipped": skipped, "scan_incomplete": scanned < len(uids) or skipped > 0,
                "results": messages}
    except MailError:
        raise
    except (imaplib.IMAP4.error, OSError, TimeoutError, ValueError):
        # Server errors can contain addresses or connection details. Return a bounded
        # generic error instead of logging IMAP command arguments/account credentials.
        raise MailError("mail_connection_failed", "无法读取校园邮箱，请核对网络、官方IMAP服务和授权状态。") from None
    finally:
        if client is not None:
            try:
                client.logout()
            except (imaplib.IMAP4.error, OSError):
                pass


def run_mail(args):
    account = args.account
    if not account:
        if not sys.stdin.isatty():
            raise MailError("mail_credentials_required", "非交互执行请指定本人邮箱 --account，密码通过授权的标准输入提供。")
        sys.stderr.write("江南大学本人邮箱地址: ")
        sys.stderr.flush()
        account = input().strip()
    if args.password_stdin:
        password = sys.stdin.readline().rstrip("\r\n")
    else:
        if not sys.stdin.isatty():
            raise MailError("mail_credentials_required", "请在交互终端输入密码，或显式使用 --password-stdin。")
        password = getpass.getpass("邮箱/客户端专用密码（不回显）: ")
    return read_mail(account, password, args.days, args.limit, args.scan_limit,
                     args.keyword, args.preview_chars)



def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--daemon-url", default="http://127.0.0.1:10086")
    parser.add_argument("--url", help="CG 课程列表链接")
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("login-status", help="检查 Edge 中 CG 课程列表是否可用")
    commands.add_parser("assignments", help="逐门核对当前作业和题目提交状态")
    mail_parser = commands.add_parser("mail", help="只读检索本人校园邮箱收件箱")
    mail_parser.add_argument("--account", help="本人完整校园邮箱地址；交互模式可省略")
    mail_parser.add_argument("--password-stdin", action="store_true", help="从授权标准输入读取密码，不接受命令行密码")
    mail_parser.add_argument("--days", type=int, default=7, help="最近多少自然日，含当天")
    mail_parser.add_argument("--limit", type=int, default=20, help="最多返回多少封")
    mail_parser.add_argument("--scan-limit", type=int, default=150, help="最多扫描多少封邮件头")
    mail_parser.add_argument("--keyword", default="", help="主题/发件人本地匹配关键字")
    mail_parser.add_argument("--preview-chars", type=int, default=0, help="可选正文文本预览字符数（0只读邮件头）")
    args = parser.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    try:
        if args.command == "mail":
            data = run_mail(args)
        else:
            url = args.url or default_url()
            base = urljoin(url, "./")
            bridge = Bridge(args.daemon_url)
            state = open_courses(bridge, url, base)
            data = ({"logged_in": True, "course_count": state["count"]}
                    if args.command == "login-status" else scan(bridge, base))
        print(json.dumps({"ok": True, "data": data}, ensure_ascii=False))
        return 0
    except (BridgeError, MailError, URLError, TimeoutError, OSError, ValueError) as error:
        code = error.code if isinstance(error, (BridgeError, MailError)) else "bridge_request_failed"
        print(json.dumps({"ok": False, "error": {"code": code, "message": str(error)}},
                         ensure_ascii=False))
        return 1


if __name__ == "__main__":
    staged = staged_command()
    if staged and staged.is_file():
        env = os.environ.copy()
        if not env.get("CG_WORKSPACE_URL") and "--url" not in sys.argv:
            try:
                env["CG_WORKSPACE_URL"] = default_url()
            except BridgeError:
                pass
        sys.exit(subprocess.call([sys.executable, str(staged), *sys.argv[1:]], env=env))
    try:
        result = main()
    finally:
        schedule()
    sys.exit(result)
