# -*- coding: utf-8 -*-
"""分级阈值校准：no7z 官方等级 vs 我们的分值等级（V43 校准第 2 步）。

数据源：资料/金标源/hsk_sentences.json（no7z/hsk-sentences-audio，4354 句，
每句 hsk_level=1..6 官方等级 + grammar_tags 官方语法点）。
本脚本：
1) 每句实跑 core.analyze_foreign_ltp → score/level；
2) 报告：官方级 vs 我们级 一致性（总体 + 每官方级混淆矩阵）；
3) 阈值扫描：在现阈值 1.4/1.75/2.0/2.3/2.55/2.7 附近 ±0.15 网格扫描，
   寻找使一致率最高的边界（≈"最接近官方分级"的映射）；
4) 输出 分析报告/校准/分级阈值校准.json + .csv（含当前与最优阈值对比）。

用法：python calibrate_levels.py [--limit N]
"""
import json
import os
import sys
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import core  # noqa: E402

SENT_PATH = os.path.join(HERE, "资料", "金标源", "hsk_sentences.json")
OUT_DIR = os.path.join(HERE, "分析报告", "校准")

CUR_BOUNDS = [1.4, 1.75, 2.0, 2.3, 2.55, 2.7]  # 当前 gf0025_level 边界


def level_for_score(score, bounds):
    for i, b in enumerate(bounds):
        if score < b:
            return i + 1
    return 7


def main():
    limit = 10 ** 9
    if "--limit" in sys.argv:
        limit = int(sys.argv[sys.argv.index("--limit") + 1])
    os.makedirs(OUT_DIR, exist_ok=True)
    sents = json.load(open(SENT_PATH, encoding="utf-8"))
    pairs = []  # (our_score, official_level)
    for s in sents[:limit]:
        text = (s.get("chinese") or "").strip()
        if not text:
            continue
        try:
            r = core.analyze_foreign_ltp(text)
        except Exception:
            continue
        pairs.append((r["score"], int(s.get("hsk_level", 0))))
    print(f"分析完成 {len(pairs)} 句")
    n = len(pairs)
    conf = Counter((p[1], level_for_score(p[0], CUR_BOUNDS)) for p in pairs)
    agree = sum(c for k, c in conf.items() if k[0] == k[1])
    print(f"当前阈值一致率: {agree}/{n} = {agree/n*100:.2f}%")
    print("混淆(官方级→我们级):")
    for (o, m), c in sorted(conf.items()):
        print(f"  官方{o}→我们{m}: {c}")
    # 阈值扫描：逐边界微调（±0.15 网格 0.05 步）以最大化一致率
    best = (agree, list(CUR_BOUNDS))
    grids = [[b + d for d in (-0.15, -0.10, -0.05, 0, 0.05, 0.10, 0.15)] for b in CUR_BOUNDS]
    import itertools
    for combo in itertools.product(*grids):
        if any(combo[i] >= combo[i + 1] for i in range(len(combo) - 1)):
            continue
        a = sum(1 for sc, lv in pairs if level_for_score(sc, combo) == lv)
        if a > best[0]:
            best = (a, list(combo))
    print(f"最优阈值一致率: {best[0]}/{n} = {best[0]/n*100:.2f}% | 边界: {best[1]}")
    # 输出
    out = {
        "sentence_count": n,
        "current_bounds": CUR_BOUNDS,
        "current_agreement": round(agree / n * 100, 2) if n else 0,
        "best_bounds": best[1],
        "best_agreement": round(best[0] / n * 100, 2) if n else 0,
        "confusion": {f"{o}-{m}": c for (o, m), c in sorted(conf.items())},
    }
    with open(os.path.join(OUT_DIR, "分级阈值校准.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    with open(os.path.join(OUT_DIR, "分级阈值校准.csv"), "w", encoding="utf-8-sig", newline="") as f:
        import csv
        w = csv.writer(f)
        w.writerow(["指标", "当前", "最优"])
        w.writerow(["边界", CUR_BOUNDS, best[1]])
        w.writerow(["一致率%", round(agree / n * 100, 2), round(best[0] / n * 100, 2)])
        w.writerow(["句子数", n, n])


if __name__ == "__main__":
    main()
