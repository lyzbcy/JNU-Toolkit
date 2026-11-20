"""Inspect CG homework through an authenticated Edge tab."""
import argparse
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


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--daemon-url", default="http://127.0.0.1:10086")
    parser.add_argument("--url", help="CG 课程列表链接")
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("login-status", help="检查 Edge 中 CG 课程列表是否可用")
    commands.add_parser("assignments", help="逐门核对当前作业和题目提交状态")
    args = parser.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    try:
        url = args.url or default_url()
        base = urljoin(url, "./")
        bridge = Bridge(args.daemon_url)
        state = open_courses(bridge, url, base)
        data = ({"logged_in": True, "course_count": state["count"]}
                if args.command == "login-status" else scan(bridge, base))
        print(json.dumps({"ok": True, "data": data}, ensure_ascii=False))
        return 0
    except (BridgeError, URLError, TimeoutError, OSError, ValueError) as error:
        code = error.code if isinstance(error, BridgeError) else "bridge_request_failed"
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
