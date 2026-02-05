import tempfile
import unittest
import zipfile
from pathlib import Path

from word_upload_demo import (
    ParagraphBlock,
    extract_docx_paragraphs,
    is_allowed_word_file,
    paginate_blocks,
)


class TestWordFileValidation(unittest.TestCase):
    def test_accept_doc(self):
        self.assertTrue(is_allowed_word_file("demo.doc"))

    def test_accept_docx_case_insensitive(self):
        self.assertTrue(is_allowed_word_file("REPORT.DOCX"))

    def test_reject_other_extensions(self):
        self.assertFalse(is_allowed_word_file("image.png"))
        self.assertFalse(is_allowed_word_file("notes.txt"))


class TestPagination(unittest.TestCase):
    def test_heading_splits_pages(self):
        blocks = [
            ParagraphBlock("封面", is_heading=True),
            ParagraphBlock("第一页正文"),
            ParagraphBlock("第二章", is_heading=True),
            ParagraphBlock("第二页正文"),
        ]
        pages = paginate_blocks(blocks, max_chars_per_page=100)
        self.assertEqual(2, len(pages))
        self.assertEqual("封面", pages[0]["title"])
        self.assertEqual("第二章", pages[1]["title"])

    def test_char_limit_splits_pages(self):
        blocks = [
            ParagraphBlock("标题", is_heading=True),
            ParagraphBlock("A" * 10),
            ParagraphBlock("B" * 10),
        ]
        pages = paginate_blocks(blocks, max_chars_per_page=15)
        self.assertEqual(2, len(pages))


class TestDocxExtraction(unittest.TestCase):
    def test_extract_docx_paragraphs_detect_heading(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "sample.docx"
            xml = """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>章节一</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>正文段落</w:t></w:r></w:p>
  </w:body>
</w:document>"""
            with zipfile.ZipFile(p, "w") as zf:
                zf.writestr("word/document.xml", xml)

            blocks = extract_docx_paragraphs(p)
            self.assertEqual(2, len(blocks))
            self.assertTrue(blocks[0].is_heading)
            self.assertEqual("正文段落", blocks[1].text)


if __name__ == "__main__":
    unittest.main()
