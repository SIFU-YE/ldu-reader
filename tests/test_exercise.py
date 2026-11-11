#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""练习起点识别逻辑测试：batch_ldu.is_exercise_start。
说明：离线主程序 batch_report_v4 已改为按 docx 全文分析，不再排除练习/正文。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx_io import is_exercise_start  # noqa: E402


# ---------------------------------------------------------------------------
# is_exercise_start：练习起点识别（用于 batch_ldu 本地分课模式）
# ---------------------------------------------------------------------------
def test_exercise_titles_detected():
    """带"练习"标题的册（第1/3/4/5册）"""
    for text in ("练习", "练习：", "练习一", "练习（一）"):
        assert is_exercise_start(text), f"应为练习起点: {text!r}"


def test_exercise_section_titles_detected():
    """无"练习"标题、直接以练习小标题开头的册（第2册）及全部练习节形式"""
    for text in (
        "一、根据文章判断正误。",
        "二、选词填空。",
        "三、根据文章回答问题。",
        "三、根据文章回答以下问题。",
        "一、根据文章选择正确答案。",
        "二、根据文章为以下事件排序。",
        "二、根据例子填空。",
        "三、讨论下列问题。",
        "一、排序。",
    ):
        assert is_exercise_start(text), f"应为练习起点: {text!r}"


def test_body_paragraphs_not_detected():
    """正文/标题段落不应被误判为练习起点"""
    for text in (
        "第1课 鱼快乐吗？",
        "第2课",
        "庄子是中国古代著名的哲学家，他和惠施是很好的朋友。",
        "一天，他们在外面散步，经过一条河，河上有一座桥。",
        "（上海一家超市的架子上放满了商品。）",
        "记者：“我是电视台记者，想问问您春节怎么过？”",
        "练习是学习的一部分。",  # 长句含"练习"字样不应误判
    ):
        assert not is_exercise_start(text), f"不应判为练习起点: {text!r}"
