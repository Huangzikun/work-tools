"""
obe_cp_sign_v6 单元测试

验证：
1. split_doc_by_page_breaks 对无分页符、显式分页符、pageBreakBefore 的正确切分
2. insert_review_to_page_center 在中间页正确插入 <wp:anchor> 浮动元素
3. 首页和末页不被插入批改记录

运行：python -m pytest obe_cp_sign/tests/test_v6_pagination.py -v
或：  python obe_cp_sign/tests/test_v6_pagination.py
"""
import os
import sys
import unittest

# 将父目录加入 sys.path 以便导入 obe_cp_sign_v6 模块
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 规避 volcenginesdkarkruntime 和 requests 依赖（单元测试不涉及网络调用）
import types as _types
if 'volcenginesdkarkruntime' not in sys.modules:
    _mock = _types.ModuleType('volcenginesdkarkruntime')
    _mock.Ark = object  # 占位类
    sys.modules['volcenginesdkarkruntime'] = _mock

from docx import Document
from docx.shared import Cm
from docx.oxml import parse_xml
from docx.oxml.ns import qn

# 导入 v6 模块（注意：v6 在模块顶层会执行 argparse，需以 importlib 加载方式规避）
import importlib.util
_v6_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "obe_cp_sign_v6.py")
_spec = importlib.util.spec_from_file_location("obe_cp_sign_v6", _v6_path)
v6 = importlib.util.module_from_spec(_spec)
# 直接加载，避免触发 argparse：手动 exec 模块代码会触发主流程
# 改为只导入函数：通过临时屏蔽 sys.argv 实现
import sys as _sys
_orig_argv = _sys.argv
_sys.argv = ["obe_cp_sign_v6"]  # 占位，避免 argparse 报错
try:
    _spec.loader.exec_module(v6)
finally:
    _sys.argv = _orig_argv


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _make_doc_with_n_pages(num_pages: int) -> Document:
    """构造一个 docx，含 num_pages 页，页之间用 <w:br w:type='page'/> 分隔"""
    doc = Document()
    for i in range(num_pages):
        p = doc.add_paragraph(f"第 {i + 1} 页内容")
        if i < num_pages - 1:
            # 在段尾添加显式分页符
            run = p.add_run()
            br = parse_xml(f'<w:br {v6._V6_NS_MAP.split(" ", 1)[0]} w:type="page"/>')
            run._r.append(br)
    return doc


def _make_doc_with_page_break_before() -> Document:
    """构造一个含 pageBreakBefore 段落属性的 docx"""
    doc = Document()
    doc.add_paragraph("第一页内容")
    p2 = doc.add_paragraph("第二页内容")
    # 给 p2 添加 pageBreakBefore 属性
    pPr = p2._p.get_or_add_pPr()
    pbb = parse_xml(f'<w:pageBreakBefore {v6._V6_NS_MAP.split(" ", 1)[0]}/>')
    pPr.append(pbb)
    doc.add_paragraph("第二页后续内容")
    return doc


class SplitDocByPageBreaksTest(unittest.TestCase):

    def test_single_page_no_break(self):
        doc = _make_doc_with_n_pages(1)
        pages = v6.split_doc_by_page_breaks(doc)
        self.assertEqual(len(pages), 1)

    def test_three_pages_explicit_break(self):
        doc = _make_doc_with_n_pages(3)
        pages = v6.split_doc_by_page_breaks(doc)
        self.assertEqual(len(pages), 3)
        # 第一页应只含 1 个 body 子元素（第一段）
        self.assertEqual(len(pages[0]), 1)
        # 第二页同样
        self.assertEqual(len(pages[1]), 1)
        self.assertEqual(len(pages[2]), 1)

    def test_page_break_before(self):
        doc = _make_doc_with_page_break_before()
        pages = v6.split_doc_by_page_breaks(doc)
        # 应切分为 2 页：第一段 + 后续两段
        self.assertEqual(len(pages), 2)


class InsertReviewToPageCenterTest(unittest.TestCase):

    def test_skip_when_less_than_three_pages(self):
        doc = _make_doc_with_n_pages(2)
        pages = v6.split_doc_by_page_breaks(doc)
        n = v6.insert_review_to_page_center(doc, pages, "评语", 80)
        self.assertEqual(n, 0)

    def test_insert_into_middle_pages(self):
        doc = _make_doc_with_n_pages(4)
        pages = v6.split_doc_by_page_breaks(doc)
        self.assertEqual(len(pages), 4)

        n = v6.insert_review_to_page_center(doc, pages, "实验完成较好", 90)
        # 中间页 = pages[1], pages[2]，共 2 页
        self.assertEqual(n, 2)

        # 检查第一页首段不含 anchor
        first_page_first_p = v6._find_first_paragraph_element(pages[0])
        anchors_in_first = first_page_first_p.findall(qn("w:drawing")) if first_page_first_p is not None else []
        # drawing 不直接是 p 的子元素，而是 run 的子元素；改为检查嵌套
        first_has_anchor = len(list(first_page_first_p.iter(qn("wp:anchor")))) > 0
        self.assertFalse(first_has_anchor, "首页不应有批改记录")

        # 检查中间页（pages[1]）含 anchor
        middle1_p = v6._find_first_paragraph_element(pages[1])
        middle1_has_anchor = len(list(middle1_p.iter(qn("wp:anchor")))) > 0
        self.assertTrue(middle1_has_anchor, "中间页 1 应有批改记录")

        # 检查中间页（pages[2]）含 anchor
        middle2_p = v6._find_first_paragraph_element(pages[2])
        middle2_has_anchor = len(list(middle2_p.iter(qn("wp:anchor")))) > 0
        self.assertTrue(middle2_has_anchor, "中间页 2 应有批改记录")

        # 检查末页不含 anchor
        last_p = v6._find_first_paragraph_element(pages[-1])
        last_has_anchor = len(list(last_p.iter(qn("wp:anchor")))) > 0
        self.assertFalse(last_has_anchor, "末页不应有批改记录")

    def test_anchor_text_contains_checkmark_and_score(self):
        doc = _make_doc_with_n_pages(3)
        pages = v6.split_doc_by_page_breaks(doc)
        v6.insert_review_to_page_center(doc, pages, "完成度好", 85)

        middle_p = v6._find_first_paragraph_element(pages[1])
        # 遍历 anchor 内的 <w:t>，断言文本含 ✓ 和 85分
        texts = [t.text for t in middle_p.iter(qn("w:t"))]
        joined = "".join(t for t in texts if t)
        self.assertIn("✓", joined)
        self.assertIn("85分", joined)
        self.assertIn("完成度好", joined)

    def test_save_and_reload(self):
        """插入后保存并重新加载，验证 anchor 持久化"""
        import tempfile
        doc = _make_doc_with_n_pages(3)
        pages = v6.split_doc_by_page_breaks(doc)
        v6.insert_review_to_page_center(doc, pages, "测试", 70)

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            doc.save(tmp_path)
            reloaded = Document(tmp_path)
            # 找到所有 anchor 元素，应至少 1 个
            anchors = list(reloaded.element.body.iter(qn("wp:anchor")))
            self.assertGreaterEqual(len(anchors), 1, "重新加载后应至少含 1 个 anchor")
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)


class BuildReviewAnchorXmlTest(unittest.TestCase):

    def test_xml_contains_position_and_text(self):
        xml = v6.build_review_anchor_xml(
            comment="评语", score=80, idx=1,
            pos_x=1000, pos_y=2000,
            box_w=3000, box_h=4000,
            font="楷体", color="FF0000", sz_half_pt=28,
        )
        self.assertIn("posOffset>1000<", xml)
        self.assertIn("posOffset>2000<", xml)
        self.assertIn("✓", xml)
        self.assertIn("80分", xml)
        self.assertIn("FF0000", xml)
        self.assertIn("楷体", xml)
        self.assertIn("wp:anchor", xml)
        self.assertIn("wps:wsp", xml)

    def test_xml_escape_special_chars(self):
        xml = v6.build_review_anchor_xml(
            comment="<script>x</script>", score=80, idx=2,
            pos_x=0, pos_y=0, box_w=100, box_h=100,
            font="宋体", color="000000", sz_half_pt=24,
        )
        self.assertNotIn("<script>", xml)
        self.assertIn("&lt;script&gt;", xml)

    def test_docpr_id_unique(self):
        xml1 = v6.build_review_anchor_xml("a", 1, 1, 0, 0, 10, 10, "宋体", "000000", 24)
        xml2 = v6.build_review_anchor_xml("b", 2, 2, 0, 0, 10, 10, "宋体", "000000", 24)
        # 提取 id 值
        import re
        id1 = re.search(r'docPr id="(\d+)"', xml1).group(1)
        id2 = re.search(r'docPr id="(\d+)"', xml2).group(1)
        self.assertNotEqual(id1, id2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
