import tempfile
import unittest
import zipfile
from pathlib import Path

from word_upload_demo import (
    ParagraphBlock,
    build_redirect_location,
    detect_person_topic,
    extract_docx_paragraphs,
    is_allowed_word_file,
    is_quote_line,
    is_section_title,
    paginate_blocks,
)


class TestWordFileValidation(unittest.TestCase):
    def test_accept_doc(self):
        self.assertTrue(is_allowed_word_file("demo.doc"))

    def test_accept_docx_case_insensitive(self):
        self.assertTrue(is_allowed_word_file("REPORT.DOCX"))

    def test_reject_other_extensions(self):
        self.assertFalse(is_allowed_word_file("image.png"))


class TestSignals(unittest.TestCase):
    def test_section_title_detection(self):
        self.assertTrue(is_section_title("一、建安风骨"))
        self.assertTrue(is_section_title("建安风骨：乱世慷慨"))

    def test_person_detection(self):
        self.assertEqual("曹操", detect_person_topic("曹操作为建安文学领袖"))

    def test_quote_detection(self):
        self.assertTrue(is_quote_line("“老骥伏枥，志在千里。”"))


class TestPagination(unittest.TestCase):
    def test_rule_based_split(self):
        blocks = [
            ParagraphBlock("一、建安风骨", is_heading=True),
            ParagraphBlock("建安文学强调现实关怀。"),
            ParagraphBlock("曹操作为建安文学领袖，风格沉郁雄健。"),
            ParagraphBlock("“老骥伏枥，志在千里。”"),
            ParagraphBlock("曹丕是曹操之子，推动七言诗成熟。"),
        ]
        pages = paginate_blocks(blocks)
        page_types = [p["page_type"] for p in pages]
        self.assertIn("section_cover", page_types)
        self.assertIn("person_profile", page_types)
        self.assertIn("quote", page_types)
        self.assertTrue(all("evidence" in p for p in pages))


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


class TestRedirectEncoding(unittest.TestCase):
    def test_redirect_location_is_ascii_and_round_trippable(self):
        location = build_redirect_location("处理完成：2 个文件", "latest_result.json")
        location.encode("latin-1")
        self.assertIn("result=latest_result.json", location)


if __name__ == "__main__":
    unittest.main()
