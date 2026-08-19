import unittest
from datetime import date, timedelta

from validation import is_valid_taiwan_id, validate_form


class ValidationTests(unittest.TestCase):
    def test_taiwan_id_checksum(self):
        self.assertTrue(is_valid_taiwan_id("A123456789"))
        self.assertFalse(is_valid_taiwan_id("A123456788"))
        self.assertFalse(is_valid_taiwan_id("123"))

    def test_valid_form(self):
        data = {
            "姓名": "測試者", "電話": "0912345678", "電子郵件": "test@example.com",
            "身分證字號": "A123456789", "目的地區": "區域", "目的監理所": "監理站",
        }
        errors = validate_form(data, date(2000, 1, 1), date.today() + timedelta(days=1))
        self.assertEqual(errors, [])

    def test_rejects_past_exam_and_bad_contact_data(self):
        data = {
            "姓名": "測試者", "電話": "abc", "電子郵件": "invalid",
            "身分證字號": "A123456788", "目的地區": "區域", "目的監理所": "監理站",
        }
        errors = validate_form(data, date(2000, 1, 1), date.today() - timedelta(days=1))
        self.assertGreaterEqual(len(errors), 4)


if __name__ == "__main__":
    unittest.main()
