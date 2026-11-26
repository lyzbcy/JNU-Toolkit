import os, sys, unittest
from pathlib import Path
from datetime import date
from email.message import EmailMessage
import imaplib

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/"scripts"))
import cg

def header(subject, sender="Example <sender@example.test>"):
    m=EmailMessage();m["Subject"]=subject;m["From"]=sender;m["Date"]="Tue, 17 Nov 2026 12:00:00 +0800"
    return m.as_bytes()

class FakeClient:
    instances=[]
    fail_login=False
    def __init__(self, host, **kwargs):
        self.host=host;self.kwargs=kwargs;self.commands=[];self.logged_out=False
        self.instances.append(self)
    def login(self, account, password):
        if self.fail_login:raise imaplib.IMAP4.error("private service response")
        self.commands.append(("LOGIN",))
    def select(self, mailbox, readonly=False):
        self.commands.append(("SELECT",mailbox,readonly))
        assert readonly is True
        return ("OK",[b"3"])
    def uid(self, operation, *args):
        self.commands.append((operation,*args))
        if operation=="SEARCH":return "OK",[b"1 2 3"]
        assert operation=="FETCH"
        uid, query=args
        assert "BODY.PEEK" in query
        if "HEADER.FIELDS" in query:
            raw={b"1":header("通知甲"),b"2":header("其他消息"),b"3":header("通知乙")}[uid]
        else:
            assert query=="(BODY.PEEK[]<0.65536>)"
            raw=b"Content-Type: text/html; charset=utf-8\r\n\r\n<p>Hello <b>world</b></p><script>not included</script>"
        return "OK",[(b"response",raw),b")"]
    def logout(self):self.logged_out=True

class MailTests(unittest.TestCase):
    def test_read_only_keyword_and_preview(self):
        result=cg.read_mail("student@stu.jiangnan.edu.cn","synthetic",keyword="通知",preview_chars=50,
                            client_factory=FakeClient)
        self.assertEqual([m["uid"] for m in result["results"]],["3","1"])
        self.assertEqual(result["results"][0]["preview"],"Hello world")
        self.assertEqual(result["scanned"],3);self.assertFalse(result["scan_incomplete"])
        client=FakeClient.instances[-1]
        self.assertTrue(client.logged_out)
        self.assertEqual(client.host,"imap.exmail.qq.com")
        self.assertTrue(client.kwargs["ssl_context"].check_hostname)
        self.assertEqual(client.commands[1],("SELECT","INBOX",True))
        self.assertTrue(all(cmd[0] in ("LOGIN","SELECT","SEARCH","FETCH") for cmd in client.commands))

    def test_limit_and_header_only(self):
        result=cg.read_mail("student@stu.jiangnan.edu.cn","synthetic",limit=1,scan_limit=2,client_factory=FakeClient)
        self.assertEqual(len(result["results"]),1);self.assertTrue(result["scan_incomplete"])
        self.assertEqual(result["scanned"],1)
        self.assertTrue(all("HEADER.FIELDS" in cmd[2] for cmd in FakeClient.instances[-1].commands if cmd[0]=="FETCH"))
        self.assertNotIn("preview",result["results"][0])

    def test_bad_account_range_and_login_not_leaking(self):
        with self.assertRaises(cg.MailError):cg.read_mail("other@example.test","synthetic",client_factory=FakeClient)
        with self.assertRaises(cg.MailError):cg.read_mail("student@stu.jiangnan.edu.cn","synthetic",days=0,client_factory=FakeClient)
        FakeClient.fail_login=True
        try:
            with self.assertRaises(cg.MailError) as cm:
                cg.read_mail("student@stu.jiangnan.edu.cn","synthetic",client_factory=FakeClient)
            self.assertEqual(cm.exception.code,"mail_login_failed")
            self.assertNotIn("private service response",str(cm.exception))
            self.assertTrue(FakeClient.instances[-1].logged_out)
        finally:FakeClient.fail_login=False

    def test_date_is_locale_independent(self):
        self.assertEqual(cg.mail_since(3,date(2026,1,2)),"31-Dec-2025")

if __name__=="__main__":unittest.main()

