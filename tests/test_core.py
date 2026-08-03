import unittest
from unittest.mock import patch

from cdr_core.jobs import get_assigned_engineer_ids, validate_job_action
from cdr_core.tokens import create_expiring_token, verify_expiring_token


class TokenTests(unittest.TestCase):
    def test_token_is_signed_and_expires(self):
        with patch("cdr_core.tokens.time.time", return_value=1_000):
            token = create_expiring_token("secret", "signature:CDR-1", lifetime_seconds=60)
        with patch("cdr_core.tokens.time.time", return_value=1_030):
            self.assertTrue(verify_expiring_token(token, "secret", "signature:CDR-1"))
            self.assertFalse(verify_expiring_token(token, "wrong", "signature:CDR-1"))
        with patch("cdr_core.tokens.time.time", return_value=1_061):
            self.assertFalse(verify_expiring_token(token, "secret", "signature:CDR-1"))


class JobRuleTests(unittest.TestCase):
    def test_engineer_ids_are_deduplicated(self):
        fields = {
            "Engineer": [{"LookupId": "7"}, {"LookupId": "7"}, {"LookupId": "8"}],
            "EngineerLookupId": "9",
        }
        self.assertEqual(get_assigned_engineer_ids(fields), ["7", "8", "9"])

    def test_job_actions_must_follow_safe_order(self):
        fields = {"Status": "Assigned", "EngineerVisitLog": ""}
        self.assertFalse(validate_job_action(fields, "Alex", "On Site")[0])
        fields["EngineerVisitLog"] = "01/01/2026 - Alex - Travelling"
        self.assertTrue(validate_job_action(fields, "Alex", "On Site")[0])


if __name__ == "__main__":
    unittest.main()
