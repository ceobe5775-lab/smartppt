import tempfile
import unittest
import zipfile
from pathlib import Path

from word_upload_demo import (
    ParagraphBlock,
    build_metadata,
    build_redirect_location,
    build_report,
    detect_person_topic,
    extract_docx_paragraphs,
    is_allowed_word_file,
    is_quote_line,
    is_section_title,
    paginate_blocks,
    split_to_bullets,
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


class TestBulletSplit(unittest.TestCase):
    def test_split_long_sentence_into_multiple_bullets(self):
        long_text = "曹操作为建安文学领袖，诗歌强调现实关怀，风格沉郁雄健，同时把个人命运与时代命运结合，形成慷慨悲凉的建安风骨，并且在语言表达上追求古朴苍凉与情感张力，使作品能够同时呈现历史厚重感与个人精神力量。"
        bullets = split_to_bullets(long_text)
        self.assertGreaterEqual(len(bullets), 2)
        self.assertTrue(all(b for b in bullets))


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

    def test_section_with_colon_has_summary_bullet(self):
        pages = paginate_blocks([ParagraphBlock("建安风骨：乱世慷慨歌壮志", is_heading=True)])
        self.assertEqual("section_cover", pages[0]["page_type"])
        self.assertEqual("建安风骨", pages[0]["title"])
        self.assertIn("乱世慷慨歌壮志", pages[0]["bullets"])


class TestGoldenQuality(unittest.TestCase):
    def test_golden_pagination_constraints(self):
        blocks = [
            ParagraphBlock("一、建安风骨", is_heading=True),
            ParagraphBlock("建安时期战乱频仍，诗歌呈现慷慨悲凉。"),
            ParagraphBlock("曹操作为建安文学领袖，强调现实关怀。"),
            ParagraphBlock("“白骨露于野，千里无鸡鸣。”"),
            ParagraphBlock("二、正始玄音", is_heading=True),
            ParagraphBlock("正始诗歌更重哲理思辨，表达幽微情感。"),
        ]
        pages = paginate_blocks(blocks)
        self.assertGreaterEqual(len(pages), 4)
        self.assertTrue(all(p["char_count"] <= 260 for p in pages))
        required_fields = {"page_type", "topic", "bullets", "quotes", "evidence", "quality_score"}
        for page in pages:
            self.assertTrue(required_fields.issubset(page.keys()))
        avg_score = sum(p["quality_score"] for p in pages) / len(pages)
        self.assertGreaterEqual(avg_score, 75)


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


class TestMetadataAndReport(unittest.TestCase):
    def test_metadata_contains_version_sha_time(self):
        metadata = build_metadata()
        self.assertIn("engine_version", metadata)
        self.assertIn("git_sha", metadata)
        self.assertIn("build_time", metadata)

    def test_build_report_contains_summary(self):
        metadata = {"engine_version": "v2", "git_sha": "abc123", "build_time": "2026-01-01T00:00:00+00:00"}
        report = build_report([{"status": "ok", "total_pages": 3, "avg_score": 90}], metadata)
        self.assertIn("engine_version: v2", report)
        self.assertIn("status_counts", report)
        self.assertIn("avg_score: 90.0", report)


class TestRedirectEncoding(unittest.TestCase):
    def test_redirect_location_is_ascii_and_round_trippable(self):
        location = build_redirect_location("处理完成：2 个文件", "latest_result.json")
        location.encode("latin-1")
        self.assertIn("result=latest_result.json", location)


if __name__ == "__main__":
    unittest.main()
