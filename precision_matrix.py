# -*- coding: utf-8 -*-
"""精确率看门狗：225 课命中的【待人工抽查表】+【跨级同形冲突】矩阵（V49 校准）。

背景：V43-V47 只用"官方例句召回率"（84.4%）衡量规则，缺精确率；且 V44 扩词
造成同形跨级重复（为什么/它们/呢/还是/怎么）。本脚本：
1) 对 语料（已拆分）/*.docx（225 课）逐课 analyze，收集【可信命中】grammar_events；
2) 输出：
   - 命中_按编号汇总.csv：编号/名称/命中次/去重句数/前5高频句（人工判断该编号精确率）
   - 命中_待人工抽查.csv：每编号 5 条高频句 + 5 条"疑似噪声句"（短/含疑问/含"怎么呢还"）
   - 跨级同形冲突.csv：同一 匹配片段（词）命中 >1 个（词类/代词/副词/助词类）编号 →
     量化"同形跨级"（A1 返工依据）
用法：python precision_matrix.py
"""
import csv
import os
import re
import sys
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import docx_io  # noqa: E402
import core  # noqa: E402

CORPUS = os.path.join(HERE, "语料（已拆分）")
OUT_DIR = os.path.join(HERE, "分析报告", "校准")


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    files = sorted(f for f in os.listdir(CORPUS) if f.endswith(".docx"))
    print("文件数:", len(files))
    by_no = defaultdict(list)         # no -> [ (fname, sent, span_text) ]
    no_hits = Counter()               # no -> 总命中次
    span_to_nos = defaultdict(set)    # 规范化词串 -> {命中过的编号}
    lesson = 0
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
        lesson += 1
        for e in r.get("grammar_events") or []:
            if e.get("status") != "trusted":
                continue
            no = e.get("no")
            span = (e.get("matched_span") or "").strip()
            sent = (e.get("sentence_id") or "")
            if not no or not span:
                continue
            no_hits[no] += 1
            by_no[no].append((fn, sent, span[:40]))
            # 词形规范化：取匹配片段去标点
            key = re.sub(r"[，。！？；：、（）\s]", "", span)[:12]
            if key:
                span_to_nos[key].add(no)
    print(f"分析 {lesson} 课，可信命中总数 {sum(no_hits.values())}")

    # 1) 按编号汇总 + 高频句
    with open(os.path.join(OUT_DIR, "命中_按编号汇总.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["编号", "命中次", "去重句数", "高频句Top5"])
        for no, cnt in no_hits.most_common():
            sents = Counter()
            for _fn, _sid, span in by_no[no]:
                sents[span] += 1
            top = [s for s, _ in sents.most_common(5)]
            w.writerow([no, cnt, len(sents), " / ".join(top)])
    # 2) 待人工抽查（每编号 5 高频 + 5 疑似噪声）
    SUSPECT = re.compile(r"[怎么呢还满都了着过是在]")
    with open(os.path.join(OUT_DIR, "命中_待人工抽查.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["编号", "类型", "句子"])
        for no in sorted(no_hits, key=lambda x: int(x)):
            recs = by_no[no]
            sents = [s for _f, _s, s in recs]
            cnt = Counter(sents)
            for s in [x for x, _ in cnt.most_common(5)]:
                w.writerow([no, "高频", s])
            for s in [x for x in sents if SUSPECT.search(x)][:5]:
                w.writerow([no, "疑似噪声", s])
    # 3) 跨级同形冲突
    with open(os.path.join(OUT_DIR, "跨级同形冲突.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["词形", "命中编号数", "编号列表"])
        for key, nos in span_to_nos.items():
            if len(nos) >= 2:
                w.writerow([key, len(nos), ",".join(sorted(nos, key=lambda x: int(x)))])
    print("输出到", OUT_DIR)


if __name__ == "__main__":
    main()
