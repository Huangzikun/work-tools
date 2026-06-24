import os
import tempfile
from pathlib import Path

import pytest

from conftest import ROSTER_FIXTURE
from services.obe_mkdir_service import (
    ObeMkdirError,
    Student,
    build_dir_tree,
    derive_major_name,
    parse_roster,
)


# ---------- derive_major_name ----------

def test_derive_major_name_with_grade_suffix():
    assert derive_major_name("2022级数据科学与大数据技术2班") == "数据科学与大数据技术2班"


def test_derive_major_name_without_grade_suffix():
    # 无 "数字+级" 前缀，原样返回
    assert derive_major_name("数据科学与大数据技术2班") == "数据科学与大数据技术2班"


def test_derive_major_name_only_digits_no_ji():
    # "2022数据科学..." 没有 "级"，不匹配，原样返回
    assert derive_major_name("2022数据科学与大数据技术2班") == "2022数据科学与大数据技术2班"


# ---------- parse_roster ----------

def test_parse_roster_from_fixture():
    data = ROSTER_FIXTURE.read_bytes()
    students = parse_roster(data, "test_roster.xls")
    assert len(students) == 5
    assert students[0].student_id == "202213008727"
    assert students[0].student_name == "王博"
    assert students[-1].student_id == "202213008801"
    assert students[-1].student_name == "李四"


def test_parse_roster_empty_table_raises():
    # 一个完全没有数字学号的名单 → 应该抛 ObeMkdirError
    html = """<html><body>
<table><tr><th>x</th></tr></table>
<table>
<thead>
<tr><th>序号</th><th>行政班级</th><th>学号</th><th>姓名</th></tr>
<tr><th>基本信息</th><th>基本信息</th><th>基本信息</th><th>基本信息</th></tr>
</thead>
<tbody>
<tr><td>1</td><td>班级A</td><td>NOT_A_NUMBER</td><td>张三</td></tr>
</tbody>
</table>
</body></html>"""
    with pytest.raises(ObeMkdirError, match="无有效学生"):
        parse_roster(html.encode("utf-8"), "empty.xls")


def test_parse_roster_missing_column_raises():
    html = """<html><body>
<table><tr><th>x</th></tr></table>
<table>
<thead>
<tr><th>序号</th><th>姓名</th></tr>
<tr><th>基本信息</th><th>基本信息</th></tr>
</thead>
<tbody>
<tr><td>1</td><td>张三</td></tr>
</tbody>
</table>
</body></html>"""
    with pytest.raises(ObeMkdirError, match="缺少必要列"):
        parse_roster(html.encode("utf-8"), "missing.xls")


# ---------- build_dir_tree ----------

def _fake_students():
    return [
        Student(seq="1", student_class="2022级xxx", student_id="202213008727", student_name="王博"),
        Student(seq="2", student_class="2022级xxx", student_id="202213008756", student_name="蒙怡冰"),
        Student(seq="3", student_class="2022级xxx", student_id="202313008747", student_name="刘悦童"),
    ]


def _walk_dirs(root):
    result = set()
    for dirpath, dirnames, _ in os.walk(root):
        for d in dirnames:
            full = os.path.join(dirpath, d)
            result.add(os.path.relpath(full, root))
    return result


def test_build_dir_tree_full_structure():
    students = _fake_students()
    with tempfile.TemporaryDirectory(prefix="obe_test_") as root:
        build_dir_tree(
            root=root,
            class_name="2022级数据科学与大数据技术2班",
            course_name="面向对象程序设计",
            teacher_name="黄子坤",
            fixed_dir_types=["教学课件", "教学教案"],
            student_dir_types=["课程考核", "实验实训报告"],
            students=students,
        )
        expected = {
            # 固定目录（不带人数）
            "2022级数据科学与大数据技术2班《面向对象程序设计》教学课件黄子坤",
            "2022级数据科学与大数据技术2班《面向对象程序设计》教学教案黄子坤",
            # 考核目录（带 3 份）
            "2022级数据科学与大数据技术2班《面向对象程序设计》课程考核黄子坤3份",
            "2022级数据科学与大数据技术2班《面向对象程序设计》实验实训报告黄子坤3份",
            # 学生子目录
            "2022级数据科学与大数据技术2班《面向对象程序设计》课程考核黄子坤3份/202213008727数据科学与大数据技术2班王博",
            "2022级数据科学与大数据技术2班《面向对象程序设计》课程考核黄子坤3份/202213008756数据科学与大数据技术2班蒙怡冰",
            "2022级数据科学与大数据技术2班《面向对象程序设计》课程考核黄子坤3份/202313008747数据科学与大数据技术2班刘悦童",
            "2022级数据科学与大数据技术2班《面向对象程序设计》实验实训报告黄子坤3份/202213008727数据科学与大数据技术2班王博",
            "2022级数据科学与大数据技术2班《面向对象程序设计》实验实训报告黄子坤3份/202213008756数据科学与大数据技术2班蒙怡冰",
            "2022级数据科学与大数据技术2班《面向对象程序设计》实验实训报告黄子坤3份/202313008747数据科学与大数据技术2班刘悦童",
        }
        actual = _walk_dirs(root)
        assert actual == expected, f"missing: {expected - actual}, extra: {actual - expected}"


def test_build_dir_tree_only_fixed():
    students = _fake_students()
    with tempfile.TemporaryDirectory(prefix="obe_test_") as root:
        build_dir_tree(
            root=root,
            class_name="2022级xxx班",
            course_name="c1",
            teacher_name="t1",
            fixed_dir_types=["教学课件"],
            student_dir_types=[],
            students=students,
        )
        dirs = _walk_dirs(root)
        assert dirs == {"2022级xxx班《c1》教学课件t1"}


def test_build_dir_tree_only_student():
    students = _fake_students()
    with tempfile.TemporaryDirectory(prefix="obe_test_") as root:
        build_dir_tree(
            root=root,
            class_name="2022级xxx班",
            course_name="c1",
            teacher_name="t1",
            fixed_dir_types=[],
            student_dir_types=["课程考核"],
            students=students,
        )
        dirs = _walk_dirs(root)
        assert "2022级xxx班《c1》课程考核t13份" in dirs
        # 学生子目录也应该在
        assert any("课程考核t13份/202213008727" in d for d in dirs)


def test_build_dir_tree_path_traversal_sanitized():
    # 学生名含路径分隔符，应该被 os.path.basename 兜底
    evil = Student(seq="1", student_class="x", student_id="123", student_name="../../etc/passwd")
    with tempfile.TemporaryDirectory(prefix="obe_test_") as root:
        build_dir_tree(
            root=root,
            class_name="班级",
            course_name="课程",
            teacher_name="教师",
            fixed_dir_types=[],
            student_dir_types=["考核"],
            students=[evil],
        )
        dirs = _walk_dirs(root)
        # 没有任何目录应该逃出 root
        for d in dirs:
            assert not d.startswith(".."), f"路径穿越！{d}"
        # 文件名被 basename 处理，应该叫 passwd
        assert any("passwd" in d for d in dirs)
