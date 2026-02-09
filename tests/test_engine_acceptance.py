import unittest

import yaml

from engine import paginate_and_classify


class TestEngineAcceptance(unittest.TestCase):
    def setUp(self) -> None:
        with open("rules.yaml", "r", encoding="utf-8") as f:
            self.rules = yaml.safe_load(f)

    def test_char_limit_and_layout(self) -> None:
        text = "\n".join(
            [
                "大家好，欢迎来到今天的课程。",
                "一、函数单调性的定义",
                "函数在某区间内若自变量增大，函数值也增大，则称在该区间单调递增。",
                "判断步骤：一是求导；二是判断导数符号；三是下结论。",
                "“老骥伏枥，志在千里；烈士暮年，壮心不已”。",
                "二、典型例题",
                "例题：已知 f(x)=x^2-2x，求其单调区间，并说明依据。",
                "希望大家课后完成配套练习。",
            ]
        )

        result = paginate_and_classify(text, self.rules)
        pages = result["pages"]
        self.assertGreaterEqual(len(pages), 2)

        max_chars = int(self.rules["engine"]["max_chars_per_page"])
        for p in pages:
            self.assertLessEqual(
                p["char_count"],
                max_chars,
                msg=f"Page {p.get('page_no')} exceeds char limit: {p}",
            )
            self.assertIn("layout", p)
            self.assertIn(
                p["layout"],
                ["全屏", "小头像", "半屏", "老师出镜", "标题页", "章节页"],
            )

        section_pages = [p for p in pages if p["page_type"] == "section_page"]
        for sp in section_pages:
            self.assertEqual(sp["layout"], "章节页")

    def test_topic_diverge_split(self) -> None:
        text = "\n".join(
            [
                "一、概念",
                "函数单调性描述函数值随自变量变化的趋势。",
                "导数符号决定单调区间。",
                "突然聊一下：唐代诗人喜欢写月亮与乡愁，这和导数没关系。",
                "回到数学：当 f'(x)>0 时函数递增。",
            ]
        )

        result = paginate_and_classify(text, self.rules)
        pages = result["pages"]
        self.assertGreaterEqual(len(pages), 2)


if __name__ == "__main__":
    unittest.main()


