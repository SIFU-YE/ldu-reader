#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""验证：本地 GF0025 公式与历史参考值的一致性"""
import csv
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import core

DATA_DIR = os.path.join(HERE, "data")


def load_levels(csv_path, level_col, key_col):
    """加载字/词分级表 -> {key: 1..7}（同一 key 多等级时取最低级，与生产 load_gf0025 口径一致）"""
    levels = {}
    with open(csv_path, encoding="utf-8") as f:
        reader = csv.reader(f)
        next(reader)
        for row in reader:
            if len(row) < max(level_col, key_col) + 1:
                continue
            key = row[key_col].strip()
            lv = row[level_col].strip()
            for prefix, n in (("一", 1), ("二", 2), ("三", 3), ("四", 4), ("五", 5), ("六", 6)):
                if lv.startswith(prefix):
                    if key not in levels or n < levels[key]:
                        levels[key] = n
                    break
            else:
                if lv.startswith("高等"):
                    if key not in levels or 7 < levels[key]:
                        levels[key] = 7
    return levels


def load_grammar(csv_path):
    """加载语法表 -> [{level, regex, project, category}]"""
    rules = []
    with open(csv_path, encoding="utf-8") as f:
        reader = csv.reader(f)
        header = next(reader)
        idx = {name: i for i, name in enumerate(header)}
        for row in reader:
            if len(row) <= idx["正则表达式"]:
                continue
            regex = row[idx["正则表达式"]].strip()
            if not regex:
                continue
            lv_text = row[idx["级别"]].strip()
            for prefix, n in (("一", 1), ("二", 2), ("三", 3), ("四", 4), ("五", 5), ("六", 6)):
                if lv_text.startswith(prefix):
                    level = n
                    break
            else:
                level = 7 if lv_text.startswith("高等") else 0
            rules.append({
                "level": level,
                "regex": regex,
                "project": row[idx["语法项目"]].strip(),
                "category": row[idx["类别"]].strip(),
            })
    return rules


POS_MAP = {
    "副词": "d", "介词": "p", "助词": "u", "量词": "q", "代词": "r",
    "动词": "v", "连词": "c", "叹词": "e", "数词": "m", "形容词": "a", "名词": "n",
}


def gf0025_score(text, char_levels, word_levels, grammar_rules):
    """程勇 GF0025 公式: 文本难度 = (1/3) * Σ_{n=1..7} n * (r_char + r_word + r_grammar)

    【V42 更正】已核对论文原文（资料/chengyong_2023.pdf）：
      n 级字比例 r_char = 课文中 n 级【字种数】/ 课文中总字种数；
      n 级词比例 r_word = 课文中 n 级【词种数】/ 课文中总词种数；
      n 级语法比例 r_grammar = 课文中 n 级【语法点类型数】/ 总语法点类型数。
    即 r 一律为【种类数/类型数】比例 —— 与生产 core.analyze_foreign_ltp 完全一致。
    本脚本历史实现（出现频次比例 + 词形单键 + 标点分句不分块）与论文不符
    （实测 1.2955 vs 1.3106，差异即由此而来），已废弃。
    现直接委托生产入口，保证“公式=论文=生产”单一口径。
    """
    import core as _core
    return _core.analyze_foreign_ltp(text)["score"]


def level_of(score):
    """分级映射，直接复用 core.gf0025_level，避免口径漂移。"""
    return core.gf0025_level(score)


if __name__ == "__main__":
    char_levels = load_levels(os.path.join(DATA_DIR, "GF0025_汉字.csv"), 1, 2)
    # 统一口径：与生产运行时一致，使用“表面形式别名表”（含 11,129 行/展开词形），
    # 而非旧运行时表 GF0025_词汇.csv（10,978 条/口径不一致）。
    word_file = os.path.join(DATA_DIR, "GF0025_词汇_别名.csv")
    if not os.path.exists(word_file):
        word_file = os.path.join(DATA_DIR, "GF0025_词汇.csv")
    word_levels = load_levels(word_file, 1, 3)
    grammar_rules = load_grammar(os.path.join(DATA_DIR, "GF0025_语法_with_regex.csv"))
    print(f"数据加载: 汉字{len(char_levels)} 词汇{len(word_levels)} 语法{len(grammar_rules)}")
    print(f"词表来源: {os.path.basename(word_file)}")

    tests = [
        ("购物短文", "我昨天去商店买东西。商店里有很多人。我买了一件衣服和一双鞋。"),
        ("课文短文", "小明今天去学校上学。他看见了一只小猫。小猫很可爱。老师教我们读书写字。我们要好好学习天天向上。"),
    ]
    for name, t in tests:
        s = gf0025_score(t, char_levels, word_levels, grammar_rules)
        print(f"{name}: 本地GF0025分值={s:.4f} -> {level_of(s)}  (论文/网站参考值: 1.2154 -> 一级难度)")
