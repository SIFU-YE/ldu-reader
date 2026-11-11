# -*- coding: utf-8 -*-
"""官方例句复验：rule 命中率/漏报清单（V43 校准第 1 步）。

数据源：资料/金标源/官方语法_例句.tsv（chelsea6502/hsk-2025-data，源 chinesetest.cn，
593 语法点 / 2055 官方例句，HSK1-7-9）。本脚本：
1) 官方语法点 → 我们 572 编号 映射（词集 Jaccard ≥0.6 记为高置信；输出映射表）；
2) 对每个官方例句实跑 core.analyze_foreign_ltp，统计映射编号的
   {官方例句数, 程序命中数, 漏报数, 命中率}；
3) 输出"漏报清单"（前 N 条：编号/官方例句）——用于人工/后续修正则；
4) 输出"候选误报"粗筛：例句未映射任何官方点，却触发某可信编号的次数（Top 编号）。

用法：python validate_against_official.py [--limit 句数上限] [--out 分析报告/校准]
输出：分析报告/校准/官方映射表.csv、官方复验_命中率.csv、官方复验_漏报清单.csv、
      官方复验_候选误报Top.csv
"""
import csv
import os
import re
import sys
from collections import defaultdict, Counter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import core  # noqa: E402

GOLD_TSV = os.path.join(HERE, "资料", "金标源", "官方语法_例句.tsv")
OUT_DIR = os.path.join(HERE, "分析报告", "校准")


def lv2num(t):
    for p, n in [("一", 1), ("二", 2), ("三", 3), ("四", 4), ("五", 5), ("六", 6)]:
        if t.startswith(p):
            return n
    return 7


def wordsets(s):
    return {p for p in re.split(r"[、,，|/;；（）()\s]+", s or "") if p and len(p) <= 4}


def build_ours():
    rows = list(csv.reader(open(os.path.join(HERE, "data", "GF0025_语法_with_regex.csv"),
                               encoding="utf-8-sig")))
    h = rows[0]
    oi = {n: i for i, n in enumerate(h)}
    by_lv = defaultdict(list)
    ours = []
    for r in rows[1:]:
        if len(r) <= oi["级别"]:
            continue
        o = {"no": r[0].strip(), "lv": lv2num(r[oi["级别"]].strip()),
             "cat": r[oi["类别"]].strip(), "det": r[oi["细目"]].strip(),
             "content": r[oi["语法内容"]].strip()}
        ours.append(o)
        by_lv[o["lv"]].append(o)
    return ours, by_lv


def make_mapping(gold_rows, by_lv, min_j=0.6):
    """官方点 → 我们编号（词集 Jaccard ≥min_j 为高置信）。
    官方点唯一键 = (examLevelId, content)（content 为行内语法点内容）。
    """
    mapping = {}
    for r in gold_rows:
        gid = r["examLevelId"]
        glv = {"HSK1": 1, "HSK2": 2, "HSK3": 3, "HSK4": 4, "HSK5": 5,
               "HSK6": 6, "HSK7-9": 7}[gid]
        wg = wordsets(r["content"]) | wordsets(r["grammarDetail"]) | wordsets(r["grammarType"])
        wg -= {"词类", "句型", "句子", "固定", "短语", "格式", "口语", "复句", "单句"}
        best, bestj = None, 0.0
        for o in by_lv.get(glv, []):
            wo = wordsets(o["content"]) | wordsets(o["det"]) | wordsets(o["cat"])
            if not wg or not wo:
                continue
            j = len(wg & wo) / len(wg | wo)
            if j > bestj:
                best, bestj = o, j
        if best and bestj >= min_j:
            mapping[(gid, r["content"], best["no"])] = bestj
    return mapping


def main():
    ap = None
    limit = 10 ** 9
    if "--limit" in sys.argv:
        limit = int(sys.argv[sys.argv.index("--limit") + 1])
    os.makedirs(OUT_DIR, exist_ok=True)
    rows = []
    with open(GOLD_TSV, encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter="\t")
        for r in reader:
            rows.append(r)
    _ours, by_lv = build_ours()
    mapping = make_mapping(rows, by_lv)
    # 1) 映射表
    with open(os.path.join(OUT_DIR, "官方映射表.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["官方ID", "等级", "官方语法点", "映射编号", "Jaccard", "官方例句数"])
        for (gid, content, no), j in sorted(mapping.items(), key=lambda kv: kv[0][0]):
            ex = sum(len([c for c in r["cases"].split("|") if c.strip()])
                     for r in rows if r["examLevelId"] == gid and r["content"] == content)
            w.writerow([gid, "", content, no, round(j, 3), ex])
    # 2) 逐句复验（targeted = 该官方点（等级+内容）映射到的编号）
    stats = defaultdict(lambda: {"ex": 0, "hit": 0})
    miss = []          # (编号, 官方例句)
    hit_by_no = Counter()  # 任意例句触发的编号计数（候选误报粗筛）
    for r in rows:
        for sent in [c.strip() for c in r["cases"].split("|") if c.strip()]:
            sent = sent[:300]
            res = core.analyze_foreign_ltp(sent)
            triggered = {n for lv in res["grammar_triggered"] for n in lv}
            targeted = [no for (gid, content, no) in mapping
                        if gid == r["examLevelId"] and content == r["content"]]
            for no in targeted:
                stats[no]["ex"] += 1
                if no in triggered:
                    stats[no]["hit"] += 1
                else:
                    miss.append((no, sent))
            for no in triggered:
                hit_by_no[no] += 1
            if len(miss) > limit * 100:
                break
        if len(miss) > limit * 100:
            break
    # 3) 命中率
    with open(os.path.join(OUT_DIR, "官方复验_命中率.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["编号", "官方例句数", "程序命中", "漏报", "命中率"])
        for no, s in sorted(stats.items(), key=lambda kv: -kv[1]["ex"]):
            if s["ex"]:
                w.writerow([no, s["ex"], s["hit"], s["ex"] - s["hit"],
                            f"{s['hit']/s['ex']*100:.1f}%"])
    # 4) 漏报清单
    with open(os.path.join(OUT_DIR, "官方复验_漏报清单.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["编号", "官方例句"])
        for no, sent in miss[:1000]:
            w.writerow([no, sent])
    # 5) 候选误报 Top（未映射点命中数）
    with open(os.path.join(OUT_DIR, "官方复验_候选误报Top.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["编号", "例句命中次数"])
        for no, c in hit_by_no.most_common(40):
            w.writerow([no, c])
    print(f"映射(≥0.6): {len(mapping)} | 复验例句: {sum(s['ex'] for s in stats.values())} | "
          f"漏报: {len(miss)}")
    print("输出:", os.path.join(OUT_DIR, "官方复验_命中率.csv"))


if __name__ == "__main__":
    main()
