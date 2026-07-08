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


if __name__ == "__main__":
    unittest.main()
