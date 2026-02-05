import unittest

from word_upload_demo import is_allowed_word_file


class TestWordFileValidation(unittest.TestCase):
    def test_accept_doc(self):
        self.assertTrue(is_allowed_word_file("demo.doc"))

    def test_accept_docx_case_insensitive(self):
        self.assertTrue(is_allowed_word_file("REPORT.DOCX"))

    def test_reject_other_extensions(self):
        self.assertFalse(is_allowed_word_file("image.png"))
        self.assertFalse(is_allowed_word_file("notes.txt"))


if __name__ == "__main__":
    unittest.main()
