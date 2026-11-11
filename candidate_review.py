# -*- coding: utf-8 -*-
"""候选回溯复核表：对已降候选的高频规则（V50），抽样其候选命中句供人工判定
是否可恢复可信（或确认降级正确）。

输出：分析报告/校准/候选_复核表.csv
    编号 / 等级 / 名称 / 候选命中次 / 高频句Top10 / 疑似噪声句Top10
用法：python candidate_review.py
"""
import csv
import os
import re
import sys
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import docx_io  # noqa: E402
import core  # noqa: E402

CORPUS = os.path.join(HERE, "语料（已拆分）")
OUT_DIR = os.path.join(HERE, "分析报告", "校准")
# V50 降候选 + 108（专项，本次一并复核）
TARGET = {"97", "98", "108", "117", "129", "130", "210"}
SUSPECT_CHARS = ("老", "小", "第", "上", "来", "起", "去", "过",
                 "一", "就", "要", "什", "么", "了", "啊", "的")


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    files = sorted(f for f in os.listdir(CORPUS) if f.endswith(".docx"))
    by_no = {}
    for fn in files:
        path = os.path.join(CORPUS, fn)
        try:
            text = "\n".join(docx_io.extract_docx_paragraphs(path)).strip()
        except Exception:
            continue
        if not text:
            continue
        try:
            r = core.analyze_foreign_ltp(text)
        except Exception:
            continue
        for e in r.get("grammar_events") or []:
            no = e.get("no")
            span = (e.get("matched_span") or "").strip()
            if not no or no not in TARGET or not span:
                continue
            by_no.setdefault(no, []).append(span[:40])
    # 等级/名称
    _cl, _wl, gr = core._get_gf0025()
    info = {r["no"]: r for r in gr if r["no"]}
    with open(os.path.join(OUT_DIR, "候选_复核表.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["编号", "等级", "名称", "候选命中次", "高频句Top10", "疑似噪声句Top10"])
        for no in sorted(TARGET, key=lambda x: int(x)):
            sents = by_no.get(no, [])
            cnt = Counter(sents)
            high = [s for s, _ in cnt.most_common(10)]
            noise = [s for s in sents if any(c in s for c in SUSPECT_CHARS)][:10]
            w.writerow([no,
                        info.get(no, {}).get("level", ""),
                        (info.get(no, {}).get("name", "") or "")[:20],
                        len(sents), " / ".join(high), " / ".join(noise)])
    print("候选复核表输出完成，TARGET 命中分布：")
    for no in sorted(by_no, key=lambda x: -len(by_no[x])):
        print(f"  {no}: {len(by_no[no])}")


if __name__ == "__main__":
    main()
