import unittest

from services.register import mail_provider
from services import register_service


class OutlookImportParseTests(unittest.TestCase):
    def test_custom_order_and_separator(self):
        # 用户自定义：email:password:refresh_token:client_id
        creds = mail_provider.parse_outlook_import(
            "PhillipaLanes4220@outlook.com:pwd123:M.C5_BAY.tok:8b4ba9dd-3ea5-4e5f-86f1-ddba2230dcf2",
            ["email", "password", "refresh_token", "client_id"], ":")
        self.assertEqual(len(creds), 1)
        self.assertEqual(creds[0]["refresh_token"], "M.C5_BAY.tok")
        self.assertEqual(creds[0]["client_id"], "8b4ba9dd-3ea5-4e5f-86f1-ddba2230dcf2")

    def test_canonical_still_parses(self):
        creds = mail_provider.parse_outlook_credentials("a@outlook.com----pw----CID----RTOK")
        self.assertEqual(creds[0], {"email": "a@outlook.com", "password": "pw", "client_id": "CID", "refresh_token": "RTOK"})

    def test_missing_required_field_rejected(self):
        # 顺序缺少 refresh_token -> 回退标准字段集，该行按标准无法凑齐 -> 丢弃
        self.assertEqual(mail_provider.parse_outlook_import("a@outlook.com:pw:CID", ["email", "password", "client_id"], ":"), [])

    def test_last_field_absorbs_extra_separator(self):
        creds = mail_provider.parse_outlook_import(
            "a@outlook.com|pw|CID|R|TOK", ["email", "password", "client_id", "refresh_token"], "|")
        self.assertEqual(creds[0]["refresh_token"], "R|TOK")

    def test_merge_stores_canonical_regardless_of_import_format(self):
        # 旧池是标准 ----，新导入是自定义 :，合并后落库仍是标准 ----
        old = "old@outlook.com----pw0----CID0----RTOK0"
        new = "new@outlook.com:pw1:RTOK1:CID1"
        merged = register_service._merge_outlook_pool(old, new, ["email", "password", "refresh_token", "client_id"], ":")
        self.assertIn("old@outlook.com----pw0----CID0----RTOK0", merged)
        self.assertIn("new@outlook.com----pw1----CID1----RTOK1", merged)


class WaitForCodeResilienceTests(unittest.TestCase):
    def _provider(self, wait_timeout=5.0):
        entry = {"type": "outlook_token", "enable": True, "mode": "imap", "mailboxes": []}
        conf = {"request_timeout": 30, "wait_timeout": wait_timeout, "wait_interval": 0, "user_agent": "UA", "proxy": ""}
        return mail_provider.OutlookTokenProvider(entry, conf)

    def test_transient_fetch_error_then_code(self):
        # 第一轮读邮箱抛瞬时错误（代理抖动），下一轮拿到验证码 -> 不该中断，应返回 code
        prov = self._provider()
        calls = {"n": 0}
        msg = {"provider": "outlook_token", "mailbox": "x", "message_id": "m1",
               "subject": "", "text_content": "Your verification code is 123456", "html_content": ""}

        def flaky(_mailbox):
            calls["n"] += 1
            if calls["n"] == 1:
                raise RuntimeError("curl (35) TLS WRONG_VERSION_NUMBER")
            return [msg]

        prov.fetch_recent_messages = flaky
        try:
            self.assertEqual(prov.wait_for_code({}), "123456")
            self.assertGreaterEqual(calls["n"], 2)
        finally:
            prov.close()

    def test_persistent_fetch_error_raises_with_reason(self):
        prov = self._provider(wait_timeout=0.4)

        def always_fail(_mailbox):
            raise RuntimeError("curl (35) TLS WRONG_VERSION_NUMBER")

        prov.fetch_recent_messages = always_fail
        try:
            with self.assertRaisesRegex(RuntimeError, "读取邮箱持续失败"):
                prov.wait_for_code({})
        finally:
            prov.close()


if __name__ == "__main__":
    unittest.main()
