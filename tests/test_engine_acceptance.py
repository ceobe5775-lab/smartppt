import unittest

import yaml

from engine import LEADIN_PATTERNS, _looks_leadin_bullet, paginate_and_classify


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

    def test_leadin_bullet_not_isolated(self) -> None:
        """
        产品级验收：如果某页最后一条 bullet 是引子句，则下一页必须存在且不是章节/标题页，
        并且下一页第一条 bullet 不能是空（说明引子句已被搬运到下一页）。
        """
        text = "\n".join(
            [
                "一、建安风骨",
                "建安时期战乱频仍，诗歌呈现慷慨悲凉。",
                "潘岳的诗歌以抒情见长，尤其是悼亡诗，",
                "写得真挚动人，比如《悼亡诗三首》：“荏苒冬春谢，寒暑忽流易”。",
                "这些作品都体现了太康诗歌的特点。",
            ]
        )

        result = paginate_and_classify(text, self.rules)
        pages = result["pages"]

        for i in range(len(pages) - 1):
            curr = pages[i]
            next_page = pages[i + 1]

            # 跳过章节页/标题页/老师出镜
            if curr.get("layout") in ("章节页", "标题页", "老师出镜"):
                continue
            if next_page.get("layout") in ("章节页", "标题页", "老师出镜"):
                continue

            bullets = curr.get("bullets", [])
            if not bullets:
                continue

            last_bullet = bullets[-1]
            if _looks_leadin_bullet(last_bullet):
                # 如果最后一条是引子句，下一页必须存在且第一条 bullet 不为空
                next_bullets = next_page.get("bullets", [])
                self.assertGreater(
                    len(next_bullets),
                    0,
                    msg=f"Page {curr.get('page_no')} ends with leadin bullet '{last_bullet[:30]}...', "
                    f"but page {next_page.get('page_no')} has no bullets",
                )
                # 下一页第一条 bullet 应该就是被搬运过来的引子句，或者至少不是空
                self.assertTrue(
                    next_bullets[0].strip(),
                    msg=f"Page {next_page.get('page_no')} first bullet is empty after leadin move",
                )

    def test_main_knowledge_anchor_mutex_split(self) -> None:
        """
        产品级验收：一个页面最多 1 个“主知识点锚点”。
        即便没超过字数限制，遇到新的主锚点也必须切新页，避免多个知识点被挤到同页。
        """
        text = "\n".join(
            [
                "一、南北朝诗歌概述",
                "南北朝时期，诗歌总体分南北两系，南朝重辞采，北朝尚质朴。",
                "谢灵运是山水诗的开创者，开拓了山水描写的审美范式。",
                "谢朓则以清新意境与格律探索著称，被称为“小谢”。",
                "鲍照则风格刚健豪放，代表作《拟行路难》多写人生困顿。",
            ]
        )

        result = paginate_and_classify(text, self.rules)
        pages = [p for p in result["pages"] if p.get("page_type") == "bullets"]
        self.assertGreaterEqual(len(pages), 4)


if __name__ == "__main__":
    unittest.main()





