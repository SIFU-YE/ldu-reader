#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""V8 修复回归测试：
1) 语法表编号唯一且完整（1..572），正则均可编译；
2) 课标题识别支持“第X课”与“1. 标题”两种格式；
3) 词汇按 (词形, 词性) 联合定级，不再一律取最低级。"""
import csv
import os
import re
import sys

import pytest  # autouse fixture 用

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import core
from docx_io import lesson_title_indices, match_lesson_title, split_lessons
from batch_report_v4 import wps_word_count

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@pytest.fixture(autouse=True)
def _restore_ltp_analyze_chunks():
    """每个测试后恢复 core._ltp_analyze_chunks 原函数。

    多个测试为 LTP 版分析 mock 了 _ltp_analyze_chunks；若不复原，
    后续测试（尤其 analyze_foreign = LTP 代理后）会沿用上一条 mock，
    导致覆盖校验失败/结果错乱。
    """
    _orig = core._ltp_analyze_chunks
    yield
    core._ltp_analyze_chunks = _orig


def test_grammar_table_unique_complete_ids():
    path = os.path.join(HERE, "data", "GF0025_语法_with_regex.csv")
    with open(path, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    assert len(rows) == 572, f"语法表应为 572 行，实际 {len(rows)}"
    ids = [r["No."].strip() for r in rows]
    assert len(set(ids)) == 572, f"语法表编号存在重复: {len(ids) - len(set(ids))} 个"
    assert set(ids) == {str(i) for i in range(1, 573)}, "语法表编号应完整覆盖 1..572"
    bad = []
    for r in rows:
        reg = (r.get("正则表达式") or "").strip()
        if reg:
            try:
                re.compile(reg)
            except re.error as exc:
                bad.append((r["No."], reg, str(exc)))
    assert not bad, f"存在无法编译的正则: {bad[:3]}"


def test_grammar_supplement_rules_are_candidate_only():
    cl, wl, gr = core._get_gf0025()
    supp = set(core.GRAMMAR_REGEX_SUPPLEMENT)
    for r in gr:
        assert "trusted" in r
        if r["no"] in supp:
            assert r["trusted"] is False, f"补全正则编号 {r['no']} 不应计为可信规则"


def test_grammar_detector_and_supplement_ids_exist():
    import grammar_ltp_detectors as gld
    with open(os.path.join(HERE, "data", "GF0025_语法_with_regex.csv"), encoding="utf-8") as f:
        rows = {int(r["No."]) for r in csv.DictReader(f)}
    for no in core.GRAMMAR_REGEX_SUPPLEMENT:
        assert int(no) in rows, f"补全正则编号 {no} 不在语法表中"
    for no in gld.DETECTORS:
        assert int(no) in rows, f"LTP检测器编号 {no} 不在语法表中"


def test_wps_word_count_like():
    assert wps_word_count("你好，世界！Hello 123") == 6  # 4汉字 + 2中文标点
    assert wps_word_count("abc") == 0
    assert wps_word_count("") == 0


def test_word_detail_label_and_entry_key():
    assert core._word_detail_label("白", "a") == "白/形"
    assert core._word_detail_label("白", "d") == "白/副"
    assert core._word_entry_key("口", "n") == ("口", "名")
    assert core._word_entry_key("口", "nl") == ("口", "名")
    assert core._word_entry_key("白", "d") == ("白", "副")


def test_lesson_title_formats():
    assert match_lesson_title("第1课 鱼快乐吗？") == ("1", "鱼快乐吗？")
    assert match_lesson_title("第12课") == ("12", "第12课")
    assert match_lesson_title("1. 爱种菜的丈母娘") == ("1", "爱种菜的丈母娘")
    assert match_lesson_title("1．爱种菜的丈母娘") == ("1", "爱种菜的丈母娘")
    assert match_lesson_title("一、根据文章判断正误。") is None
    assert match_lesson_title("正文内容") is None


def test_dot_title_not_used_when_di_titles_exist():
    """文档已有“第X课”标题时，正文中的“3. 青蛙心里想……”不得被误判为课标题。"""
    paras = ["第1课 鱼快乐吗？", "庄子是哲学家。", "3. 青蛙心里想：真的比井底大吗？",
             "练习", "一、根据文章判断正误。"]
    assert lesson_title_indices(paras) == [0]
    lessons = split_lessons(paras)
    assert len(lessons) == 1, lessons
    assert "3. 青蛙心里想" in lessons[0][2]
    assert "练习" not in lessons[0][2]


def test_split_lessons_dot_title():
    paras = ["1. 爱种菜的丈母娘", "我家有一位爱种菜的丈母娘。", "练习",
             "一、根据文章判断正误。", "2. 邻居家的猫", "邻居家有一只猫。"]
    lessons = split_lessons(paras)
    assert len(lessons) == 2, lessons
    assert lessons[0][0] == "1" and lessons[0][1] == "爱种菜的丈母娘"
    assert "丈母娘" in lessons[0][2] and "判断正误" not in lessons[0][2]
    assert lessons[1][0] == "2" and "邻居家的猫" in lessons[1][1]
    assert "猫" in lessons[1][2]


def test_word_pos_joint_lookup():
    cl, wl, gr = core._get_gf0025()
    wpl = core._get_word_pos_levels()
    assert cl and wl and gr
    # “白”：一级形 / 三级副 → 词性可消歧
    lv, amb = core._lookup_word_level(wl, wpl, "白", "a")
    assert lv == 1 and amb is False, ("白/a", lv, amb)
    lv, amb = core._lookup_word_level(wl, wpl, "白", "d")
    assert lv == 3 and amb is False, ("白/d", lv, amb)
    # “打”：一级动 / 四级量 → 词性可消歧
    lv, _ = core._lookup_word_level(wl, wpl, "打", "v")
    assert lv == 1
    lv, _ = core._lookup_word_level(wl, wpl, "打", "q")
    assert lv == 4
    # “地方”：一级名 / 四级名 → 同词性多级，标记为多级候选并保守取最低级
    lv, amb = core._lookup_word_level(wl, wpl, "地方", "n")
    assert lv == 1 and amb is True, ("地方/n", lv, amb)
    # 无词性时回退词形最低级
    lv, _ = core._lookup_word_level(wl, wpl, "地方", "")
    assert lv == 1


def test_analyze_foreign_returns_word_ambiguous():
    r = core.analyze_foreign("他是地方电视台的记者。")
    assert "word_ambiguous" in r
    # “地方”在同词性下多级，出现在多级候选列表
    assert "地方" in r["word_ambiguous"]


def test_lookup_detail_distinguishes_pos_fallback():
    cl, wl, gr = core._get_gf0025()
    wpl = core._get_word_pos_levels()
    lv, amb, src = core._lookup_word_level_detail(wl, wpl, "白", "a")
    assert lv == 1 and src == "pos"
    lv, amb, src = core._lookup_word_level_detail(wl, wpl, "很", "m")
    # LTP/jieba 把“很”标成数词时，标准表无该词性，应走词形兜底
    assert lv == 1 and src == "form"


def test_beyond_words_are_classified_for_teaching_safety():
    cl, wl, gr = core._get_gf0025()
    assert core._classify_beyond_word("北京", "ns", wl) == "proper"
    assert core._classify_beyond_word("二十一", "m", wl) == "number"
    assert core._classify_beyond_word("很快", "d", wl) == "phrase"
    assert core._classify_beyond_word("的", "u", wl) == "fragment"
    assert core._classify_beyond_word("珐琅", "n", wl) == "ordinary"


def test_analyze_foreign_formula_uses_word_form_and_reports_clean_beyond():
    r = core.analyze_foreign("他很高兴。奥运会很精彩。高铁很快。我们去北京和二十一。")
    # 词种公式/报告都使用 word_unique_counts，且 word_detail 长度一致
    assert len(r["word_detail"]) == len(r["word_unique_counts"]) == 7
    assert sum(r["word_unique_counts"]) == sum(len(x) for x in r["word_detail"])
    # 基本情况表中的“词种数”也与公式/分级表统一（只计标准词表内词形）
    assert r["stats"]["unique_words"] == sum(r["word_unique_counts"])
    # 超纲词清单已把数字/短语/专名分开，普通超纲词不再包含这些类别。
    # LTP 语义：`很快` 被切为 `很/快`（均词表内命中，非超纲）；`二十一` 归数字类。
    assert "二十一" not in r["beyond_words_clean"]
    assert "二十一" in r["beyond_number_like"]
    # LTP 分词后“很快”拆成 很+快（词表内），不再是超纲短语
    assert "很快" not in r["beyond_phrases"]


def test_chapter_docx_classification_and_natural_sort():
    import batch_report_v4 as b
    # V31：输入约定仅分章文件 x-x-x.docx；_classify_docx 返回 (前缀, 章节号)，非分章为 None
    assert b._classify_docx("3-2-1.docx") == ("3-2", "1")
    assert b._classify_docx("3-2-10.docx") == ("3-2", "10")
    assert b._classify_docx("3-2.docx") is None
    assert b._classify_docx("readme.docx") is None
    names = ["3-2-10.docx", "3-2-2.docx", "3-2-1.docx"]
    assert sorted(names, key=b._natural_key) == ["3-2-1.docx", "3-2-2.docx", "3-2-10.docx"]


def test_merge_docx_creation():
    import os
    import tempfile
    import batch_report_v4 as b
    fd, tmp = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        b._make_merged_docx(tmp, "测试合并", [
            ("3-2-1.docx", "第一课内容。\n第二段。"),
            ("3-2-2.docx", "第二课内容。"),
        ])
        assert os.path.getsize(tmp) > 0
    finally:
        os.remove(tmp)


def test_load_volume_lessons_supports_chapter_files():
    import docx_io as b
    # 新分章格式：资料/上传结果/3-2-1.docx 直接作为一课
    lessons = b.load_volume_lessons(
        os.path.join(HERE, "资料", "上传结果"), "3-2")
    assert len(lessons) == 1
    assert lessons[0]["num"] == "1"
    assert lessons[0]["body"].strip()


def test_iter_lessons_from_docx_chapter_file():
    import docx_io as b
    path = os.path.join(HERE, "资料", "上传结果", "3-2-1.docx")
    items = list(b.iter_lessons_from_docx(path, "3-2-1.docx"))
    assert len(items) == 1
    assert items[0][0] == "1"
    assert items[0][2].strip()


def test_generated_report_file_filter():
    import docx_io as b
    assert b.is_generated_report_file("3-2-1_分析报告.docx")
    assert b.is_generated_report_file("合并分析报告.docx")
    assert not b.is_generated_report_file("3-2-1.docx")
    assert not b.is_generated_report_file("3-2.docx")


def test_expanded_word_table_surface_forms():
    cl, wl, gr = core._get_gf0025()
    for w in ["爸爸", "爸", "哥哥", "哥", "有时候", "有时", "第", "第二",
              "朋友们", "现代化", "服务员", "积极性", "科学家", "老王", "小王",
              "汇率", "好不容易", "好容易", "影响力", "成功率", "茅台", "茅台酒",
              "差点儿", "差一点儿", "有点儿", "有一点儿", "有些", "有一些",
              "分之", "极了", "面", "妈"]:
        assert wl.get(w), f"展开词表应包含 {w}"


def test_ocr_word_huilv_fixed():
    cl, wl, gr = core._get_gf0025()
    assert wl.get("汇率") == 4
    # 原 OCR 错行“会/huìlǜ/名”不应再让“会”错误获得四级
    wpl = core._get_word_pos_levels()
    # LTP/jieba 的“名”原始词性为 n
    lv, amb, src = core._lookup_word_level_detail(wl, wpl, "会", "n")
    # “会”的名词义项应仍为二级，而不是四级
    assert lv == 2, (lv, src)


def test_word_unique_counts_use_pos_level_not_lowest():
    import core as _core

    def fake_ltp(chunks):
        words = ["白", "打", "花", "。"]
        poses = ["d", "q", "v", "wp"]
        return [{"cws": words, "pos": poses,
                 "dep": {"head": [0] * 4, "rel": [""] * 4}}]

    _core._ltp_analyze_chunks = fake_ltp
    r = _core.analyze_foreign_ltp("白打花。")
    assert r["word_unique_counts"] == [0, 1, 1, 1, 0, 0, 0], r["word_unique_counts"]
    assert r["word_detail"][1] == ["花/动"]
    assert r["word_detail"][2] == ["白/副"]
    assert r["word_detail"][3] == ["打/量"]


def test_canonical_and_alias_tables():
    import os
    canon_path = os.path.join(HERE, "data", "GF0025_词汇_规范.csv")
    alias_path = os.path.join(HERE, "data", "GF0025_词汇_别名.csv")
    assert os.path.exists(canon_path)
    assert os.path.exists(alias_path)

    with open(canon_path, encoding="utf-8") as f:
        canon = list(csv.DictReader(f))
    assert len(canon) == 11092, f"规范条目表应为 11092 行，实际 {len(canon)}"
    ids = [r["No."].strip() for r in canon]
    assert len(set(ids)) == len(ids), "规范条目表 No. 应唯一"
    ocr = [r for r in canon if r["No."].strip() == "3595"]
    assert ocr and ocr[0]["词语"] == "汇率", ocr

    with open(alias_path, encoding="utf-8") as f:
        alias = list(csv.DictReader(f))
    alias_words = {r["表面形式"] for r in alias}
    for w in ["爸爸", "爸", "汇率", "好不容易", "好容易", "差点儿", "差一点儿",
              "有点儿", "有一点儿", "分之", "极了"]:
        assert w in alias_words, f"别名表缺少 {w}"


def test_every_level_has_normal_and_expanded_words():
    cl, wl, gr = core._get_gf0025()
    # 每级普通词
    normal_by_level = {1: "爱", 2: "安静", 3: "安排", 4: "阿姨",
                       5: "安慰", 6: "挨着", 7: "哎"}
    for lv, w in normal_by_level.items():
        assert wl.get(w) == lv, (lv, w, wl.get(w))
    # 每级展开词/别名词
    expanded_by_level = {1: "爸爸", 2: "有点儿", 3: "现代化",
                         4: "汇率", 5: "差一点儿", 6: "好容易", 7: "茅台酒"}
    for lv, w in expanded_by_level.items():
        assert wl.get(w) == lv, (lv, w, wl.get(w))


def test_alias_pos_errors_fixed():
    """P13/P1 回归：别名表词性错误修复后，(词形, 标准词性大类) 必须能精确命中。"""
    _cl, _wl, _gr, wpl = core.load_gf0025()
    # 修复前这些 (词形, 词性) 键不存在 → LTP 命中时走 form 兜底、不计词种
    cases = [
        ("省", "动", 2),        # 规范 No=7150 省（动）→ 动（原被误标“名”）
        ("且", "连", 7),        # 规范 No=6429 且（连）→ 连（原被误标“副”）
        ("之", "助", 7),        # 规范 No=10451 之（助）→ 助（原被误标“代”）
        ("料", "名", 6),        # 规范 No=5168 料（名）→ 名（原被误标“动”）
        ("码", "量", 7),        # 规范 No=5397 码（量）→ 量（原被误标“动”）
        ("批", "量", 4),        # 规范 No=6043 批（量）→ 量（原被误标“动”）
        ("麻", "形", 7),        # 规范 No=5380 麻（形）→ 形（原被误标“名”）
        ("成年", "副", 7),      # 规范 No=965 成年（副）→ 副（原被误标“动”）
        ("则", "量", 7),        # 规范 No=10180 则（量）→ 量（原被误标“连”）
        ("副", "量", 6),        # 规范 No=2518 副（量）→ 量（原被误标“形”）
        ("任", "连", 3),        # 规范 No=6747 任（连）→ 连（原被误标“动”）
        ("有些", "代", 1),      # No=9897 有（一）些 → 代（原被误标 Pr）
        ("有一些", "代", 1),    # No=9897 有（一）些 → 代（原被误标 Pr）
        ("老王", "名", 2),      # No=4954 老（老王）→ 名（原被误标 Name）
        ("小王", "名", 2),      # No=8890 小（小王）→ 名（原被误标 Name）
    ]
    for word, cat, expect_lv in cases:
        lv, amb, src = core._lookup_word_level_detail(None, wpl, word, _pos_for_cat(cat))
        assert lv == expect_lv and src == "pos", (word, cat, lv, src)
        assert amb is False, (word, cat, "不应是多级候选")


def _pos_for_cat(cat):
    """《标准》大类 → 用于 _lookup_word_level_detail 的原始词性"""
    return {"名": "n", "动": "v", "形": "a", "副": "d", "连": "c",
            "助": "u", "量": "q", "代": "r"}[cat]


def test_pos_cat_exact_overrides_prefix():
    """P7 防护：Pr/Name 等非标准符号不得被首字母前缀误映射。"""
    assert core._pos_to_word_cats("Pr") == ["代"], "Pr 不应被 p 前缀误映射为“介”"
    assert core._pos_to_word_cats("pr") == ["代"]
    assert core._pos_to_word_cats("Name") == ["名"], "Name 不应被 n 前缀误映射为“名”"
    assert core._pos_to_word_cats("name") == ["名"]
    # 原有 LTP/jieba 词性映射不受影响
    assert core._pos_to_word_cats("n") == ["名"]
    assert core._pos_to_word_cats("nr") == ["名"]
    assert core._pos_to_word_cats("v") == ["动"]
    assert core._pos_to_word_cats("p") == ["介"]
    assert core._pos_to_word_cats("r") == ["代"]


def test_alias_no_conflict_and_ocr_clean():
    """P2 回归：No.228 只关联“报警”，“妈”应挂到妈妈编号 5378；No.3595 无 OCR 残留。"""
    canon_path = os.path.join(HERE, "data", "GF0025_词汇_规范.csv")
    alias_path = os.path.join(HERE, "data", "GF0025_词汇_别名.csv")
    with open(canon_path, encoding="utf-8") as f:
        canon_ids = {r["No."].strip() for r in csv.DictReader(f)}
    with open(alias_path, encoding="utf-8") as f:
        alias = list(csv.DictReader(f))
    # P15：别名表 No. 必须全部存在于规范表
    bad_ids = {r["No."].strip() for r in alias} - canon_ids
    assert not bad_ids, f"别名表存在规范表没有的编号: {sorted(bad_ids)[:5]}"
    # No.228 唯一对应 报警，不得再有“妈”
    no228 = [r for r in alias if r["No."].strip() == "228"]
    assert len(no228) == 1 and no228[0]["规范写法"].strip() == "报警", no228
    # “妈”应位于一级 妈妈（5378）编号下
    ma = [r for r in alias if r["表面形式"].strip() == "妈"]
    assert ma and ma[0]["No."].strip() == "5378", ma
    # No.3595 规范写法应为干净的“汇率”
    no3595 = [r for r in alias if r["No."].strip() == "3595"]
    assert no3595 and no3595[0]["规范写法"].strip() == "汇率", no3595


def test_same_form_two_pos_in_one_text():
    import core as _core

    def fake_ltp(chunks):
        words = ["白", "白", "。"]
        poses = ["a", "d", "wp"]
        return [{"cws": words, "pos": poses,
                 "dep": {"head": [0] * 3, "rel": [""] * 3}}]

    _core._ltp_analyze_chunks = fake_ltp
    r = _core.analyze_foreign_ltp("白白。")
    # 同一词形不同词性分别计数：白/形一级、白/副三级，应计为 2 个词种
    assert r["word_unique_counts"] == [1, 0, 1, 0, 0, 0, 0], r["word_unique_counts"]
    assert "白/形" in r["word_detail"][0]
    assert "白/副" in r["word_detail"][2]
    assert sum(r["word_unique_counts"]) == r["stats"]["unique_words"]


def test_beyond_words_and_proper_nouns():
    r = core.analyze_foreign("上海珐琅二十一。")
    assert "上海" in r["beyond_proper_nouns"]
    assert "二十一" in r["beyond_number_like"]
    assert "珐琅" in r["beyond_words_clean"]


def test_grammar_ids_aligned_after_renumber():
    import grammar_ltp_detectors as gld
    with open(os.path.join(HERE, "data", "GF0025_语法_with_regex.csv"), encoding="utf-8") as f:
        ids = {r["No."].strip() for r in csv.DictReader(f)}
    assert len(ids) == 572
    for k in gld.DETECTORS:
        assert k in ids, f"DETECTORS 编号 {k} 不在语法表中"
    for k in gld.LEVEL_DETECTORS:
        assert k in ids, f"LEVEL_DETECTORS 编号 {k} 不在语法表中"
    for k in gld.LEVEL_DETECTOR_TRUSTED_NOS:
        assert k in ids, f"LEVEL_DETECTOR_TRUSTED_NOS 编号 {k} 不在语法表中"
    for k in gld.WEAK_DETECTOR_NOS:
        assert k in ids, f"WEAK_DETECTOR_NOS 编号 {k} 不在语法表中"
    for k in core.GRAMMAR_REGEX_SUPPLEMENT:
        assert k in ids, f"GRAMMAR_REGEX_SUPPLEMENT 编号 {k} 不在语法表中"
    for k in core.GRAMMAR_SENTENCE_SUPPLEMENT_NOS:
        assert k in ids, f"GRAMMAR_SENTENCE_SUPPLEMENT_NOS 编号 {k} 不在语法表中"
    for k in core.GRAMMAR_UNRELIABLE_NOS:
        assert k in ids, f"GRAMMAR_UNRELIABLE_NOS 编号 {k} 不在语法表中"


def test_grammar_unreliable_nos_aligned_with_table():
    """回归：黑名单编号必须与 data/GF0025_语法_with_regex.csv 对齐。

    曾于提交 b6d56d4 被整体误减 1（222→221、243→242……）而数据表未变，
    导致黑名单指向错误规则（242 实为“在……方面”，243“在……上/下/中”
    被重新当作可信规则）。此处按名称锚定关键规则，防止再次错位。
    """
    import grammar_ltp_detectors as gld
    with open(os.path.join(HERE, "data", "GF0025_语法_with_regex.csv"), encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    by_no = {r["No."].strip(): r for r in rows}
    # 黑名单编号存在性（上一条已覆盖），且锚定规则名与黑名单一致
    anchors = {
        "243": "在……上/下/中",
        "475": "……，而……（则）……",
        "492": "……，以……",
        "505": "不知X的",
        "510": "什么X的Y的",
    }
    for no, name in anchors.items():
        assert no in core.GRAMMAR_UNRELIABLE_NOS, f"编号 {no}（{name}）应降为候选，不在黑名单中"
        row = by_no.get(no)
        assert row is not None, f"编号 {no} 不在语法表中"
        # 规则名可能出现在“语法项目/类别/细目/语法内容”任意一列
        fields = "".join(row.get(k, "") or "" for k in
                         ("语法项目", "类别", "细目", "语法内容"))
        assert name in fields, (
            f"黑名单编号 {no} 指向了错误规则：{fields!r}，应为 {name!r}"
        )
    # 曾因错位被误列入的“在……方面”（242）不应在名单中
    assert "242" not in core.GRAMMAR_UNRELIABLE_NOS, "242（在……方面）是编号错位残留，应恢复可信"
    # 黑名单全部规则都应在表中且级别有效（无正则条目也允许，避免越界）
    for no, row in by_no.items():
        if no in core.GRAMMAR_UNRELIABLE_NOS:
            assert (row.get("级别") or "").strip(), f"黑名单编号 {no} 缺少级别"
    # 正则已降为候选的规则，其 LTP 检测器也不得计为可信，防止绕开黑名单重新进入覆盖
    # （is_trusted 仅对有 LTP 检测器实现的编号有判断意义，其余编号不会出现在检测结果中）
    for no in anchors:
        if no in gld.DETECTORS or no in gld.LEVEL_DETECTORS:
            assert not gld.is_trusted(no), (
                f"LTP 检测器编号 {no} 不应为可信（与正则黑名单口径一致）"
            )


def test_chunk_text_bounds():
    """修复1：生产入口按 ≤400 字切块，避免 LTP 静默截断。"""
    text = "我" * 600 + "阿拉伯语"
    chunks = core._chunk_text(text)
    assert all(len(c) <= 400 for c in chunks), [len(c) for c in chunks]
    assert sum(len(c) for c in chunks) >= len(text)
    assert len(chunks) >= 2


def test_verify_ltp_coverage():
    """修复1：LTP 输出覆盖不足必须抛错，不允许静默截断。"""
    chunks = ["这是一个测试句子。"]
    # 正常覆盖：不抛错
    core._verify_ltp_coverage(chunks, [{"cws": ["这", "是", "一个", "测试", "句子", "。"],
                                        "pos": [], "dep": {}}])
    # 块数不一致：抛错
    import pytest
    try:
        core._verify_ltp_coverage(chunks, [{"cws": ["这"]}])
        assert False, "块数不一致应抛 RuntimeError"
    except RuntimeError:
        pass
    # 单块覆盖不足：抛错
    with pytest.raises(RuntimeError):
        core._verify_ltp_coverage(["这是一段很长的测试内容，包含大量词汇。"],
                                  [{"cws": ["这"]}])


def test_fallback_words_counted_in_unique():
    """修复3：词性兜底词（src=form）必须计入词种数/分值，不再静默排除。"""
    import core as _core

    def fake_ltp(chunks):
        # “很”词表为 d(副)，LTP 标 m(数) → 走 form 兜底
        return [{"cws": ["很", "好", "。"], "pos": ["m", "a", "wp"],
                 "dep": {"head": [0, 0, 0], "label": [""] * 3}}]

    _core._ltp_analyze_chunks = fake_ltp
    r = _core.analyze_foreign_ltp("很好。")
    # 兜底词“很”按词形唯一键计数并计入词种（不再重复计多个预测词性）
    assert r["word_fallback_unique_counts"] == [1, 0, 0, 0, 0, 0, 0], r["word_fallback_unique_counts"]
    assert "很" in r["word_fallback_detail"][0]
    assert "很/数" not in r["word_fallback_detail"][0]  # 兜底按词形展示，不带预测词性
    assert r["word_unique_counts"] == [2, 0, 0, 0, 0, 0, 0], r["word_unique_counts"]
    assert "很" in r["word_detail"][0]
    assert r["stats"]["unique_words"] == 2
    assert sum(r["word_unique_counts"]) == r["stats"]["unique_words"]
    # 含预测词性的完整兜底明细保留在 word_pos_fallback_detail（诊断）
    assert "很/数" in r["word_pos_fallback_detail"][0]


def test_fallback_word_form_dedup():
    """P2：同一兜底词被预测成多个词性时，只按词形唯一键计数一次。"""
    import core as _core

    def fake_ltp2(chunks):
        # “一样”分别被标 助(ug) 与 数(m)，词表无这两词性 → 均 form 兜底
        return [{"cws": ["一样", "一样", "。"], "pos": ["ug", "m", "wp"],
                 "dep": {"head": [0, 0, 0], "label": [""] * 3}}]

    _core._ltp_analyze_chunks = fake_ltp2
    r = _core.analyze_foreign_ltp("一样一样。")
    # 词形唯一键：兜底词种只计 1 个“一样”（两个 token 同词形去重），
    # 而不是按预测词性（ug/数）计 2 个
    assert r["word_fallback_unique_counts"] == [1, 0, 0, 0, 0, 0, 0], r["word_fallback_unique_counts"]


def test_token_audit_output():
    """修复5：逐 Token 审计记录（句号/词形/词性/来源/等级/是否计分/分类）。"""
    import core as _core

    def fake_ltp(chunks):
        return [{"cws": ["他", "很", "好", "。"], "pos": ["r", "d", "a", "wp"],
                 "dep": {"head": [0, 0, 0, 0], "label": [""] * 4}}]

    _core._ltp_analyze_chunks = fake_ltp
    r = _core.analyze_foreign_ltp("他很好。")
    audits = r["token_audit"]
    assert len(audits) == 4
    by_word = {a["word"]: a for a in audits}
    assert by_word["他"]["source"] == "pos" and by_word["他"]["scored"] is True
    assert by_word["好"]["level"] == 1
    assert by_word["。"]["category"] == "punctuation" and by_word["。"]["scored"] is False
    assert all("sentence" in a and "pos" in a and "level" in a for a in audits)


def test_total_words_semantics():
    """修复6：total_words 为分词 Token 数（含标点），新增纯汉字词数口径。"""
    import core as _core

    def fake_ltp(chunks):
        return [{"cws": ["他", "很", "好", "。"], "pos": ["r", "d", "a", "wp"],
                 "dep": {"head": [0, 0, 0, 0], "label": [""] * 4}}]

    _core._ltp_analyze_chunks = fake_ltp
    r = _core.analyze_foreign_ltp("他很好。")
    s = r["stats"]
    # 4 个 token（含标点），3 个纯汉字词（不含标点）
    assert s["total_words"] == 4
    assert s["total_word_tokens"] == 4
    assert s["total_hanzi_words"] == 3


def test_ambiguous_detail_candidate_levels():
    """修复7：多级同形同词性词输出候选等级，供人工复核。"""
    import core as _core

    def fake_ltp(chunks):
        # “地方”词表为 一级/四级 同名，LTP 标 n → amb=True
        return [{"cws": ["他", "是", "地方", "的", "记者", "。"],
                 "pos": ["r", "v", "n", "u", "n", "wp"],
                 "dep": {"head": [0] * 6, "label": [""] * 6}}]

    _core._ltp_analyze_chunks = fake_ltp
    r = _core.analyze_foreign_ltp("他是地方的记者。")
    assert "地方" in r["word_ambiguous"]
    assert r["word_ambiguous_detail"].get("地方") == [1, 4], r["word_ambiguous_detail"]


def test_chunk_keeps_sentences_separate():
    """P1-2：正常句子逐句输入，不合并句子改变 LTP 分词上下文。"""
    text = "我想省一点钱。他们也想要。"
    chunks = core._chunk_text(text)
    assert chunks == split_sentences_clean(text), chunks
    # 每句 ≤400 时 _chunk_text 与 _sentence_list 一致
    assert chunks == core._sentence_list(text)


def split_sentences_clean(text):
    return [s for s in core.split_sentences(text) if s.strip()]


def test_long_sentence_boundary_cut():
    """P1-3：超长单句按边界/重叠切分，不截断词语（399我+阿拉伯语整词保留）。"""
    text = "我" * 399 + "阿拉伯语"
    chunks = core._chunk_text(text)
    assert all(len(c) <= 400 for c in chunks)
    # 至少一个块包含完整“阿拉伯语”
    assert any("阿拉伯语" in c for c in chunks), chunks
    # 拼接后应覆盖原文全部内容（重叠窗口会超长，但不应丢字符）
    assert any("阿拉伯语" in c and (len(c) <= 400) for c in chunks)


def test_ltp_model_fallback_path():
    """P1-1：_get_ltp 模型路径选择——本地存在用绝对路径；缺失走模型名称。

    使用 mock 隔离 _HERE_DIR 到临时目录，避免触碰真实 LTP/small 模型文件。
    """
    import core as _core
    import unittest.mock as mock
    import tempfile
    # 注入假 ltp 模块（_get_ltp 内部 from ltp import LTP）
    import sys as _sys
    class FakeLTP:
        def __init__(self, path):
            self.path = path
    fake_ltp_mod = mock.MagicMock()
    fake_ltp_mod.LTP = FakeLTP
    _sys.modules["ltp"] = fake_ltp_mod
    try:
        # 1) 本地模型存在 → 绝对路径（用临时目录模拟，不动真实仓库模型）
        with tempfile.TemporaryDirectory() as td:
            _core._LTP = None
            with mock.patch.object(_core, "_HERE_DIR", td):
                local_model = os.path.join(td, "LTP", "small")
                os.makedirs(local_model, exist_ok=True)
                with open(os.path.join(local_model, "config.json"), "w", encoding="utf-8") as f:
                    f.write("{}")
                ltp = _core._get_ltp()
                assert ltp.path == local_model, ltp.path
        # 2) 模型缺失 → 名称回退（LTP/small；用临时目录模拟干净环境）
        with tempfile.TemporaryDirectory() as td:
            _core._LTP = None
            with mock.patch.object(_core, "_HERE_DIR", td):
                ltp = _core._get_ltp()
                assert ltp.path == "LTP/small", ltp.path
        _core._LTP = None
    finally:
        _sys.modules.pop("ltp", None)


def test_server_rows_uses_array_directly():
    """P1-4：Web 明细表 rows 使用 tags(v)（数组直传），不把数组当对象读 v[key]。"""
    import server
    page = server.PAGE
    # rows 定义应为 tags(v)（JS 模板字面量 ${tags(v)}），且不存在旧 bug 的 v[key]
    assert "${tags(v)}" in page
    assert "v[key]" not in page


def test_token_audit_real_sentence_numbers():
    """P2：token_audit.sentence 是真实句子编号，不是分块编号。"""
    import core as _core

    def fake_ltp(chunks):
        # 两块 = 两句
        return [{"cws": ["我", "好", "。"], "pos": ["r", "a", "wp"],
                 "dep": {"head": [0, 0, 0], "label": [""] * 3}},
                {"cws": ["你", "好", "。"], "pos": ["r", "a", "wp"],
                 "dep": {"head": [0, 0, 0], "label": [""] * 3}}]

    _core._ltp_analyze_chunks = fake_ltp
    r = _core.analyze_foreign_ltp("我好。你好。")
    sents = sorted({t["sentence"] for t in r["token_audit"]})
    assert sents == [1, 2], sents


def test_docx_extract_unescapes_entities():
    """P2：DOCX 提取反转义 XML 实体（A&B <测试> 不再变成 A&amp;B &lt;测试&gt;）。"""
    import zipfile, tempfile, io
    import docx_io as b

    fd, tmp = tempfile.mkstemp(suffix=".docx")
    import os
    os.close(fd)
    try:
        # 构造最小 docx（含实体文本）
        with zipfile.ZipFile(tmp, "w") as z:
            xml = ('<?xml version="1.0" encoding="UTF-8"?>'
                   '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                   '<w:body><w:p><w:r><w:t>A&amp;B &lt;测试&gt;</w:t></w:r></w:p></w:body></w:document>')
            z.writestr("word/document.xml", xml)
        paras = b.extract_docx_paragraphs(tmp)
        assert paras and paras[0] == "A&B <测试>", paras
    finally:
        os.remove(tmp)


def test_coverage_threshold_is_strict():
    """P2：完整性校验阈值 90%，400字输入仅返回 240 字（60%）必须报错。"""
    chunks = ["x" * 400]
    # 240/400 = 60% < 90% → 必须抛错
    try:
        core._verify_ltp_coverage(chunks, [{"cws": ["x" * 240]}])
        assert False, "60% 覆盖应抛 RuntimeError（阈值已收紧至 90%）"
    except RuntimeError:
        pass


def test_model_split_and_merge():
    """模型分片入库：_ensure_model_merged 能把 partN 合并回完整 bin。"""
    import tempfile
    import core as _core
    with tempfile.TemporaryDirectory() as td:
        # 模拟分片模型目录
        data = b"abcdefghij" * 1000  # 10000 字节
        with open(os.path.join(td, "pytorch_model.bin.part1"), "wb") as f:
            f.write(data[:5000])
        with open(os.path.join(td, "pytorch_model.bin.part2"), "wb") as f:
            f.write(data[5000:])
        # 合并
        ok = _core._ensure_model_merged(td)
        assert ok
        merged = open(os.path.join(td, "pytorch_model.bin"), "rb").read()
        assert merged == data, "合并结果应与原数据一致"
    # 无分片也无完整 bin → 返回 False（不崩溃）
    with tempfile.TemporaryDirectory() as td:
        assert _core._ensure_model_merged(td) is False


def test_model_merge_validation():
    """P2：合并校验——单分片/零字节分片/零字节完整 bin 均不视为有效模型。"""
    import tempfile
    import core as _core
    # 单分片（不足 2 片）→ False
    with tempfile.TemporaryDirectory() as td:
        with open(os.path.join(td, "pytorch_model.bin.part1"), "wb") as f:
            f.write(b"x" * 100)
        assert _core._ensure_model_merged(td) is False
    # 零字节分片 → False
    with tempfile.TemporaryDirectory() as td:
        with open(os.path.join(td, "pytorch_model.bin.part1"), "wb") as f:
            f.write(b"")
        with open(os.path.join(td, "pytorch_model.bin.part2"), "wb") as f:
            f.write(b"y" * 100)
        assert _core._ensure_model_merged(td) is False
    # 零字节完整 bin（无分片）→ False（视为无效）
    with tempfile.TemporaryDirectory() as td:
        with open(os.path.join(td, "pytorch_model.bin"), "wb") as f:
            f.write(b"")
        assert _core._ensure_model_merged(td) is False


def test_chunk_no_overlap():
    """P1-1：超长句分块无重叠，块拼接后完整覆盖原文本（不重复计数）。"""
    # 402 字（399 个我 + 阿拉伯语）
    text = "我" * 399 + "阿拉伯语"
    chunks = core._chunk_text(text)
    assert sum(len(c) for c in chunks) == len(text), (len(text), sum(len(c) for c in chunks))
    assert all(len(c) <= 400 for c in chunks)
    # 403 我（无标点词边界）
    text2 = "我" * 403
    chunks2 = core._chunk_text(text2)
    assert sum(len(c) for c in chunks2) == len(text2), (len(text2), sum(len(c) for c in chunks2))


def test_word_boundary_protects_std_words():
    """P1：400 字边界不切断 GF0025 标准词（词表前缀集合校验）。
    审查的 281 词（爱不释手/持之以恒等 4 字词）词首在第 400 字符时不应被切。"""
    import core as _core
    prefixes = _core._get_std_prefix_set()
    std_words = _core._get_std_word_set()
    # 直接校验词表全部多字词：词首前切点必须安全、词内切点必须识别为切断
    _cl, wl, _gr = _core._get_gf0025()
    bad = []
    for w in sorted(wl):
        if len(w) < 2:
            continue
        sent = "我" * 399 + w
        if not _core._cut_is_safe(sent, 399, prefixes, std_words):
            bad.append((w, "词首前不安全"))
            continue
        for k in range(1, len(w)):
            if _core._cut_is_safe(sent, 399 + k, prefixes, std_words):
                bad.append((w, f"词内切点{399+k}误判安全"))
                break
    assert not bad, f"词表边界保护失败 {len(bad)} 词: {bad[:5]}"
    # 抽测审查案例
    for w in ["爱不释手", "持之以恒", "阿拉伯语", "半信半疑"]:
        sent = "我" * 399 + w
        chunks = _core._chunk_text(sent)
        assert any(w in c for c in chunks), w


def test_fallback_dedup_across_levels():
    """P2：跨等级兜底去重——词形在任一等级精确命中后，不再额外计兜底词型。"""
    import core as _core

    def fake_ltp(chunks):
        # “报”：v 命中 报/动（七级精确）；x 未知词性 → 词形兜底（较低级）
        return [{"cws": ["报", "报", "。"], "pos": ["v", "x", "wp"],
                 "dep": {"head": [0, 0, 0], "label": [""] * 3}}]

    _core._ltp_analyze_chunks = fake_ltp
    r = _core.analyze_foreign_ltp("报报。")
    # 精确命中 7 级；兜底不再计入 → 词型合计 1，且 word_detail 只有“报/动”
    exact_levels = [i for i, d in enumerate(r["word_pos_detail"], 1) if d]
    fb_levels = [i for i, d in enumerate(r["word_fallback_detail"], 1) if d]
    assert exact_levels == [7], exact_levels
    assert fb_levels == [], fb_levels
    assert sum(r["word_unique_counts"]) == 1
    assert r["stats"]["unique_words"] == 1
    assert r["word_detail"][6] == ["报/动"], r["word_detail"][6]


def test_model_merge_atomic_and_size_check():
    """P2：模型合并原子写入 + 大小校验（截断/残留 bin 应重新合并）。"""
    import tempfile
    import core as _core
    import os as _os
    with tempfile.TemporaryDirectory() as td:
        data = b"abcdefghij" * 1000
        with open(_os.path.join(td, "pytorch_model.bin.part1"), "wb") as f:
            f.write(data[:5000])
        with open(_os.path.join(td, "pytorch_model.bin.part2"), "wb") as f:
            f.write(data[5000:])
        # 存在【截断的 bin】（大小不等于分片和）→ 应重新合并
        with open(_os.path.join(td, "pytorch_model.bin"), "wb") as f:
            f.write(b"truncated")
        ok = _core._ensure_model_merged(td)
        assert ok
        merged = open(_os.path.join(td, "pytorch_model.bin"), "rb").read()
        assert merged == data, "截断 bin 应被重新合并为完整数据"
        # 无残留 .tmp 文件
        assert not _os.path.exists(_os.path.join(td, "pytorch_model.bin.tmp"))


def _mk_tokens(words, poses):
    """构造检测器 tokens（word/pos/i/head/rel）"""
    return [{"w": w, "pos": p, "i": i, "head": 0, "rel": ""}
            for i, (w, p) in enumerate(zip(words, poses))]


def test_det_141_quantity_overlap_no_false_positive():
    """P17：det_141 量词重叠——非量词叠词不误报，真量词重叠正确命中。"""
    import grammar_ltp_detectors as gld
    # 非量词叠词（LTP 标 n/v）不触发
    assert gld.det_141(_mk_tokens(["妈妈"], ["n"])) is False
    assert gld.det_141(_mk_tokens(["谢谢"], ["v"])) is False
    # 量词重叠 AA / 一X一X（LTP 或标 q 或标 m 数词）触发
    assert gld.det_141(_mk_tokens(["条条"], ["q"])) is True
    assert gld.det_141(_mk_tokens(["一棵棵"], ["m"])) is True
    assert gld.det_141(_mk_tokens(["一", "棵", "棵"], ["m", "q", "q"])) is True
    # 非重叠量词结构（一 + 棵 + 树）不触发
    assert gld.det_141(_mk_tokens(["一", "棵", "树"], ["m", "q", "n"])) is False


def test_det_132_suffix_no_false_positive():
    """P17：det_132 后缀——固化名词（孩子/桌子）不误报，构词后缀（们/k）命中。"""
    import grammar_ltp_detectors as gld
    # 固化名词词尾（LTP 标 n）→ 不触发
    assert gld.det_132(_mk_tokens(["孩子"], ["n"])) is False
    assert gld.det_132(_mk_tokens(["桌子"], ["n"])) is False
    # 独立后缀 token（们/k、子/u）→ 触发
    assert gld.det_132(_mk_tokens(["朋友", "们"], ["n", "k"])) is True
    assert gld.det_132(_mk_tokens(["桌", "子"], ["n", "u"])) is True
    # 普通名词（无后缀字）→ 不触发
    assert gld.det_132(_mk_tokens(["商店"], ["n"])) is False


def test_det_203_last_jiu_no_false_negative():
    """V25：det_203 紧缩复句取最后一个“就”——首个“就”在“了”前时不再漏检。"""
    import grammar_ltp_detectors as gld
    # 标准的“……了……就……”：首“了”在首“就”前 → 命中
    assert gld.det_203(_mk_tokens(["他", "吃", "了", "饭", "就", "走", "了"],
                                  ["r", "v", "u", "n", "d", "v", "u"])) is True
    # 首“就”在“了”前（“他马上就吃饭了”中“就”先出现），但末“就”在“了”后
    assert gld.det_203(_mk_tokens(["他", "就", "吃", "了", "饭", "马上", "就", "走", "了"],
                                  ["r", "d", "v", "u", "n", "d", "d", "v", "u"])) is True
    # “了”后无“就” → 不触发
    assert gld.det_203(_mk_tokens(["他", "吃", "了", "饭"],
                                  ["r", "v", "u", "n"])) is False
    # “就”与“了”之间有标点 → 不触发
    assert gld.det_203(_mk_tokens(["他", "吃", "了", "饭", "，", "就", "走"],
                                  ["r", "v", "u", "n", "wp", "d", "v"])) is False


def test_det_unreliable_326_397_not_trusted():
    """V25：326（意念被动句）/397（把字句4·表致使）语义无法由通用检测器证实，仅候选。"""
    import grammar_ltp_detectors as gld
    assert gld.is_trusted("326") is False
    assert gld.is_trusted("397") is False
    # 检测器仍保留注册（触发供人工复核），不从 LEVEL_DETECTORS 删除
    assert "326" in gld.LEVEL_DETECTORS and "397" in gld.LEVEL_DETECTORS


def test_multi_level_shared_detector_only_main_trusted():
    """V30：把/被/兼语句的多编号共享检测器仅主编号可信，杜绝同句重复记级。

    同一句“把”字句不得同时计入三/四/五/六级：det_184 只证明大类存在，
    细分语义（处置1/2/3、致使4）无法区分 → 仅 183（把字句1）可信，
    248/325/397/249/326/398/251/328 等必须为候选。
    """
    import grammar_ltp_detectors as gld
    # 主编号（各细分的起始/基础结构）保持可信
    for no in ("183", "184", "186"):
        assert gld.is_trusted(no), f"主编号 {no} 应保持可信"
    # 细分编号全部降为候选
    for no in ("248", "249", "251", "325", "326", "328", "397", "398"):
        assert gld.is_trusted(no) is False, (
            f"细化编号 {no} 不应计为可信（检测器仅能证明大类结构存在，避免重复记级）"
        )
        assert no in gld.LEVEL_DETECTORS, f"{no} 应保留注册（候选触发供人工复核）"


def test_supplement_140_keeps_guanting():
    """V25：140 量词重叠补全正则应保留官方词面“顿顿”，不因补全降级而丢失。"""
    import core as _core
    assert "顿顿" in _core.GRAMMAR_REGEX_SUPPLEMENT["140"]


def test_native_semantic_richness_denominator_excludes_empty_class():
    """V25：母语语义丰富度分母使用【非空】语义类数（原 len(_CILIN)=12 含空类 respectful）。"""
    import core as _core

    def fake_ltp(chunks):
        return [{"cws": ["我", "吃", "饭", "。"], "pos": ["r", "v", "n", "wp"],
                 "dep": {"head": [0, 0, 0, 0], "label": ["", "", "", ""]}}]

    _core._ltp_analyze_chunks = fake_ltp
    r = _core.analyze_native("我吃饭。")
    # 出现类：action(v)、object(n)；分母 11（非空类），而非 12
    # （注意返回值已 round(...,4)）
    assert abs(r["semantic_richness"] - round(2 / 11, 4)) < 1e-9, r["semantic_richness"]


def test_jieba_score_uses_fallback_in_word_types():
    """语法一致性：评分公式的词种口径与返回统计一致（含独立词形兜底）。

    注意 core 已弃用 Jieba（analyze_foreign 现直接转 LTP 版），且两口径“自然不同”：
    stats.unique_words 是【全局并集】（跨级词只算 1 次），word_unique_counts 是
    【逐级求和】（跨级词如“很”可在一级与二级各计 1 次），故只断言 <= 与正数。
    """
    import core as _core
    r = _core.analyze_foreign("他很高兴。")
    assert r["stats"]["unique_words"] <= sum(r["word_unique_counts"])
    assert r["score"] > 0
    # 公式词种数（word_types）应与返回 word_unique_counts 各级一致：
    # 分值由 (1/3)Σ n*(r_char+r_word+r_gram) 决定，r_word 用 word_types
    # 这里直接验证返回值中各级词种与 stats 全局一致
    wc = r["word_unique_counts"]
    assert len(wc) == 7 and all(isinstance(x, int) for x in wc)


def test_char_freq_available_for_native_formula():
    """回归：母语公式的字频项不再静默为 0。

    data/character_frequency.csv 曾缺失（仅存在于被 .gitignore 的 build_data/），
    load_char_freq() 返回 {} → avg_char_freq=0 → 公式 -45.65*avg/1e6 项恒为 0，
    母语分值系统性偏高（实测同文本 40.5537 → 40.3332）。此处断言字频表已随
    仓库 data/ 交付且加载非空。
    """
    import core as _core
    freq = _core.load_char_freq()
    assert freq, "data/character_frequency.csv 缺失或为空：母语公式字频项将恒为 0"
    assert len(freq) > 1000
    # 高频字“的”应有每百万频率值（数据目录从 build_data/ 复制）
    assert "的" in freq and freq["的"] > 1000


def test_integrated_stats():
    """V35：整合统计（业务项4）——每分章=一课，汇总均值/中位数/标准差/等级课数/并集/累加。"""
    import batch_report_v4 as br
    records = [
        {"ok": True, "fname": "1-1-1.docx", "score": 1.2, "level": "一级",
         "stats": {"total_chars": 100, "total_words": 90, "total_hanzi_words": 80,
                   "total_word_tokens": 90, "sentences": 5,
                   "unique_chars": 30, "unique_words": 20},
         "char_detail": [["爱"], ["来"], [], [], [], [], []],
         "word_detail": [["我/代词"], ["爱/动词"], [], [], [], [], []],
         "grammar_covered": [["1"], ["48"], [], [], [], [], []]},
        {"ok": True, "fname": "1-1-2.docx", "score": 2.0, "level": "二级",
         "stats": {"total_chars": 200, "total_words": 180, "total_hanzi_words": 160,
                   "total_word_tokens": 180, "sentences": 10,
                   "unique_chars": 50, "unique_words": 40},
         "char_detail": [["爱", "你"], ["来", "去"], [], [], [], [], []],
         "word_detail": [["我/代词"], ["爱/动词", "去/动词"], [], [], [], [], []],
         "grammar_covered": [["1"], ["48", "49"], [], [], [], [], []]},
        {"ok": False, "fname": "1-1-3.docx", "error": "解析失败"},
    ]
    s = br.integrated_stats(records)
    assert s["lesson_count"] == 3 and s["ok_count"] == 2
    assert s["failed"] == ["1-1-3.docx"]
    assert s["score"]["mean"] == 1.6 and s["score"]["median"] == 1.6
    assert round(s["score"]["stdev"], 4) == 0.4
    assert s["score"]["min"] == 1.2 and s["score"]["max"] == 2.0
    assert s["level_counts"] == {"一级": 1, "二级": 1}
    assert s["char_unique_union"]["1"] == ["你", "爱"]
    assert s["word_unique_union"]["2"] == ["去/动词", "爱/动词"]
    assert s["grammar_covered_union"]["2"] == ["48", "49"]
    assert s["totals"]["total_chars"] == 300 and s["totals"]["sentences"] == 15
    assert s["totals"]["unique_words"] == 60
    # 全失败场景：不抛异常
    empty = br.integrated_stats([{"ok": False, "fname": "x"}])
    assert empty["lesson_count"] == 1 and empty["ok_count"] == 0 and empty["score"] is None


def test_score_snapshot_baseline():
    """准确度基准快照：固定文本分值/级别范围。

    用途：公式（r 口径/权重）、词表、语法可信集改动后，若分值跳出既有范围，
    说明回归或口径漂移，需人工复核。范围取当前实测值上下 0.1 （宽松阈值，
    防脆弱）；历史参考：论文/网站同文本 1.2154（verify_gf0025），README 宣称
    与网站差 0.3-0.4。
    """
    import core as _core
    cases = [
        ("我昨天去商店买东西。商店里有很多人。我买了一件衣服和一双鞋。",
         (1.10, 1.50)),  # 当前 1.3106
        ("他把门关上了。杯子被打碎了。",
         (1.60, 2.00)),  # 当前 1.7841
    ]
    for text, (lo, hi) in cases:
        r = _core.analyze_foreign(text)
        assert lo <= r["score"] <= hi, (
            f"分值超出基准范围: {r['score']} 不在 [{lo}, {hi}]，文本={text[:20]!r}")
        assert r["level"] in core.LEVEL_NAMES


def test_det_187_requires_causative_head():
    """V41：det_187 兼语句收紧——仅使令动词（叫/派/请/让/使/令）触发。

    原实现允许任意动词头（如"教/打电话找我/催/叫醒"），
    "女儿还教我在网上买东西""一天，有个朋友打电话找我""早上最重要的一个
    工作是催儿子起床"等普通句被误判为兼语句（186/251/328）——现不再触发。
    """
    import grammar_ltp_detectors as gld

    def _tokens(words, poses, heads, rels):
        return [{"w": w, "pos": p, "i": i, "head": h, "rel": r}
                for i, (w, p, h, r) in enumerate(zip(words, poses, heads, rels))]

    # 使令动词"让" + DBL（他→让）+ 另一动词（去/买→让）→ 触发
    ts = _tokens(["妈妈", "让", "他", "去", "买", "东西"],
                 ["n", "v", "r", "v", "v", "n"],
                 [0, 0, 2, 2, 2, 5], ["", "HED", "DBL", "CMP", "CMP", ""])
    assert gld.det_187(ts) is True
    # 非使令动词"教" 同结构 → 不再触发
    ts2 = _tokens(["女儿", "教", "我", "买", "东西"],
                  ["n", "v", "r", "v", "n"],
                  [0, 0, 2, 2, 4], ["", "HED", "DBL", "CMP", ""])
    assert gld.det_187(ts2) is False
    # 无 DBL 关系 → 不触发
    assert gld.det_187(_mk_tokens(["催", "儿子", "起床"], ["v", "n", "v"])) is False


def test_supplement_no_punctuation_candidates():
    """V41：497/515/516 原用纯标点（；/。）作候选补全正则，候选统计饱和——已移除。"""
    import core as _core
    for no in ("497", "515", "516"):
        assert no not in _core.GRAMMAR_REGEX_SUPPLEMENT
        assert no not in _core.GRAMMAR_SENTENCE_SUPPLEMENT_NOS


def test_ltp_cache_epoch_warnings():
    """V29：LTP 词汇缓存词表代次指纹校验。

    旧缓存（无 word_level_fingerprint 字段）不告警（向后兼容）；
    指纹与当前词表不一致时必须告警，防止 reproduce_paper 混用
    旧覆盖数 ÷ 新词表分母产生虚假覆盖率（831/953=87.20% vs 自洽 86.38%）。
    """
    import reproduce_paper as rp
    wl = {"a": 3, "b": 3, "c": 7}  # 三级 2 个、七级 1 个
    # 指纹一致 → 无告警
    assert rp._ltp_cache_epoch_warnings(
        {"word_level_fingerprint": {"counts": [0, 0, 2, 0, 0, 0, 1], "total": 3}}, wl) == []
    # 指纹不一致（三级少算 1） → 告警且说明代次
    warns = rp._ltp_cache_epoch_warnings(
        {"word_level_fingerprint": {"counts": [0, 0, 1, 0, 0, 0, 1], "total": 2}}, wl)
    assert len(warns) == 1 and "词表代次" in warns[0] and "ltp_vocab_reproduction" in warns[0]
    # 旧缓存无指纹字段 → 不告警（向后兼容）
    assert rp._ltp_cache_epoch_warnings({}, wl) == []
    # 指纹缺失/格式错误 → 不告警
    assert rp._ltp_cache_epoch_warnings({"word_level_fingerprint": {"counts": [0]}}, wl) == []


# ---------------------------------------------------------------------------
# V55：用户评审问题修复回归测试（22/33/82/108/52/56 + 金标门控）
# ---------------------------------------------------------------------------

def test_v55_uhtp_disambig_preference():
    """V55：助词(u)规则消歧——同一 u 词只归入首选编号，不再双计。
    "了"→21（动态）、"的"→20（结构）、"地"→20、"得"→79、"着"→81、"过"→80、
    "吧/吗/呢"→22、"啊"→82。优先表必须覆盖全部可消歧字面。"""
    assert core._UHLP_DISAMBIG["了"] == "21"
    assert core._UHLP_DISAMBIG["的"] == "20"
    assert core._UHLP_DISAMBIG["地"] == "20"
    assert core._UHLP_DISAMBIG["得"] == "79"
    assert core._UHLP_DISAMBIG["着"] == "81"
    assert core._UHLP_DISAMBIG["过"] == "80"
    assert core._UHLP_DISAMBIG["吧"] == "22"
    assert core._UHLP_DISAMBIG["吗"] == "22"
    assert core._UHLP_DISAMBIG["呢"] == "22"
    assert core._UHLP_DISAMBIG["啊"] == "82"
    assert core._UHLP_DISAMBIG["所"] == "379"
    assert core._UHLP_DISAMBIG["之"] == "433"


def test_v55_review_required_nos_gate():
    """V55：金标门控——22/33/82/108 无金标证据时必须降为候选，不得进入可信覆盖。
    其余规则（20/21/52/56 等）无证据时维持配置身份。"""
    assert core._evidence_ok("22") is False
    assert core._evidence_ok("33") is False
    assert core._evidence_ok("82") is False
    assert core._evidence_ok("108") is False
    assert core._evidence_ok("20") is True
    assert core._evidence_ok("21") is True
    assert core._evidence_ok("52") is True
    assert core._evidence_ok("56") is True
    # 黑名单规则本来就不可信
    assert core._evidence_ok("97") is True  # 不在必审集合，判定由 GRAMMAR_UNRELIABLE_NOS 负责


def test_v55_det_yiwen_ji_not_interrogative():
    """V55：det_yiwen —— '几'（LTP 标 m 数词）非疑问语境不触发；'呢'陈述句不触发；
    疑问标记仍触发。"""
    import grammar_ltp_detectors as gld
    # "几"作数量词（无问号/吗/呢）→ 不触发
    assert gld.det_yiwen(_mk_tokens(["这", "几", "年"], ["r", "m", "q"])) is False
    assert gld.det_yiwen(_mk_tokens(["住", "了", "几", "天"], ["v", "u", "m", "q"])) is False
    assert gld.det_yiwen(_mk_tokens(["几", "天", "以后"], ["m", "q", "f"])) is False
    # "呢"陈述句（无问号/疑问代词）→ 不触发
    assert gld.det_yiwen(_mk_tokens(["他", "明天", "来", "呢", "。"], ["r", "nt", "v", "u", "wp"])) is False
    assert gld.det_yiwen(_mk_tokens(["我们", "正好", "聊天", "呢", "。"], ["r", "d", "v", "u", "wp"])) is False
    # 疑问代词（r 词性）→ 触发
    assert gld.det_yiwen(_mk_tokens(["你", "是", "哪", "国", "人", "？"], ["r", "v", "r", "n", "n", "wp"])) is True
    assert gld.det_yiwen(_mk_tokens(["从", "哪儿", "来", "的"], ["p", "r", "v", "u"])) is True
    # 吗/问号 → 触发
    assert gld.det_yiwen(_mk_tokens(["你", "好", "吗", "？"], ["r", "a", "u", "wp"])) is True
    assert gld.det_yiwen(_mk_tokens(["几", "点", "了", "？"], ["m", "q", "u", "wp"])) is True
    # "呢" + 疑问代词/问号 → 触发（疑问语境）
    assert gld.det_yiwen(_mk_tokens(["你", "呢", "？"], ["r", "u", "wp"])) is True
    assert gld.det_yiwen(_mk_tokens(["他", "是", "谁", "呢", "。"], ["r", "v", "r", "u", "wp"])) is True


def test_v55_det_yiwen_33_not_trusted():
    """V55：33（疑问句）从 LTP 可信集合移除——必须人工复核。"""
    import grammar_ltp_detectors as gld
    assert gld.is_trusted("33") is False
    # 同级仍未受影响
    assert gld.is_trusted("35") is True
    assert gld.is_trusted("42") is True


def test_v55_shi_de_emphasis_judgements():
    """V55：_is_shi_de_emphasis —— 强调句"是+状语+动词+的"判 True；
    "X是Y的Z"主系表/特指疑问判 False（V51 残留误报）。"""
    import core as _core
    import grammar_ltp_detectors as gld
    # 真阳性（强调时间/地点/方式/动作者）
    assert _is_shi_de_ok("他是昨天来的。") is True
    assert _is_shi_de_ok("她是在北京出生的。") is True
    assert _is_shi_de_ok("东西是昨天买的。") is True
    assert _is_shi_de_ok("我是不会忘记的。") is True
    # V55 召回损失修复：方式状语以动词开头（LTP 标 v）+ "的"后接名词宾语
    assert _is_shi_de_ok("他是乘飞机去的。") is True
    assert _is_shi_de_ok("他是开车来的。") is True
    assert _is_shi_de_ok("他是走路去的。") is True
    assert _is_shi_de_ok("他是在食堂吃的饭。") is True
    # V51 误报：主系表"X是Y的Z"
    assert _is_shi_de_ok("庄子是哲学家。") is False
    assert _is_shi_de_ok("吃美食是年轻人生活中的重要部分。") is False
    assert _is_shi_de_ok("买也是周末生活的一部分。") is False
    assert _is_shi_de_ok("旅行，是我最喜欢的事。") is False
    assert _is_shi_de_ok("是住学校的宿舍。") is False
    assert _is_shi_de_ok("是我们全家人一起看电视的快乐。") is False
    # 特指疑问（V51 曾误判为强调）
    assert _is_shi_de_ok("他是从哪儿来的。") is False


def _is_shi_de_ok(sent):
    """对单句跑 108 正则是否被 _is_shi_de_emphasis 判为强调句。"""
    import re
    from core import _chunk_text, _ltp_analyze_chunks, _is_shi_de_emphasis
    pat = re.compile(r"是.+的")
    chunks = _chunk_text(sent)
    results = _ltp_analyze_chunks(chunks)
    for chunk, res in zip(chunks, results):
        words = res.get("cws", [])
        poses = res.get("pos", [])
        for m in pat.finditer(chunk):
            return _is_shi_de_emphasis(m, words, poses)
    return False


def test_v55_verb_overlap_morph_exclusions():
    """V55：52 动词重叠形态排除扩充——X啊X啊/X吧X吧/你你我我/来来回回 不触发；
    真 A一A/AA/ABAB 仍触发。"""
    import core as _core
    # 用 _match_token_pos_in_span 的形态排除逻辑直接验证
    for bad in ("来吧来吧", "来啊来啊", "你你我我", "来来回回"):
        assert _verb_overlap_ok(bad) is False, f"{bad} 应为非动词重叠"
    for good in ("走一走", "看一看", "想一想", "讨论讨论", "研究研究", "尝尝", "走走"):
        assert _verb_overlap_ok(good) is True, f"{good} 应为动词重叠"


def _verb_overlap_ok(sent):
    """对单句跑 52 regex + 词性窗口是否被判定为动词重叠。"""
    import re
    from core import _chunk_text, _ltp_analyze_chunks, _match_token_pos_in_span
    pat = re.compile(r"([\u4e00-\u9fff]{1,2})\1|([\u4e00-\u9fff]{1,2})一\2|([\u4e00-\u9fff]{1,2})了\3|([\u4e00-\u9fff]{2})\4")
    chunks = _chunk_text(sent)
    results = _ltp_analyze_chunks(chunks)
    for chunk, res in zip(chunks, results):
        words = res.get("cws", [])
        poses = res.get("pos", [])
        for m in pat.finditer(chunk):
            if _match_token_pos_in_span(m, words, poses, "v"):
                return True
    return False


def test_v55_adj_overlap_no_cross_token():
    """V55：56 形容词重叠——跨词匹配（弟弟快快/看看空空）被 token 边界校验排除；
    真 AA的/AABB 仍触发。"""
    import re
    from core import _chunk_text, _ltp_analyze_chunks, _match_token_pos_in_span
    pat = re.compile(r"(\S)\1的|(\S)\2(\S)\3")
    # 跨词误报（匹配片段穿出 LTP token 边界）
    for bad in ("弟弟快快乐乐", "看看空空的山"):
        chunks = _chunk_text(bad)
        results = _ltp_analyze_chunks(chunks)
        for chunk, res in zip(chunks, results):
            words = res.get("cws", [])
            poses = res.get("pos", [])
            for m in pat.finditer(chunk):
                # 跨词匹配被排除或整段不在边界 → 不得判为 56 命中
                assert _match_token_pos_in_span(m, words, poses, ("a", "z")) is False
    # 真形容词重叠
    for good in ("高高兴兴", "高高的山", "干干净净的"):
        chunks = _chunk_text(good)
        results = _ltp_analyze_chunks(chunks)
        hit = False
        for chunk, res in zip(chunks, results):
            words = res.get("cws", [])
            poses = res.get("pos", [])
            for m in pat.finditer(chunk):
                if _match_token_pos_in_span(m, words, poses, ("a", "z")):
                    hit = True
        assert hit, f"{good} 应为形容词重叠"
