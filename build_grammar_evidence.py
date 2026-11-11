# -*- coding: utf-8 -*-
"""从 语法金标样本.csv 生成 语法规则证据.json（第6项样本验收门控的数据源）。

口径（《程序运行逻辑说明（修订版）》）：
- 候选规则升级为可信：至少抽查 10 个真实命中点、错误 ≤ 1（精确率 ≥ 90%）；
- 误差定义：句子被程序【触发】（正则+LTP 任一命中 = grammar_triggered）而金标=0 → 误报（错误）；
  金标=1 未触发为漏报（仅记录 fn，不计入精确率，因为验收针对"真实命中点"）；
- 当前金标样本 15 条/4 编号（183/184/186/187），每编号 3-4 条 < 10：
  按规则判定为【样本不足 → 自动降级为候选】（core._evidence_ok 返回 False）。

用法：python build_grammar_evidence.py（需 LTP 环境）；输出 data/语法规则证据.json。
      python build_grammar_evidence.py --build-pool --dir ./资料
        → 从语料抽取“官方可信句式/固定格式/口语格式”规则的命中句，输出
          data/语法金标候选池.csv（编号/等级/名称/句子/程序触发/金标【留空待人工补标】），
          供人工标注后并入 语法金标样本.csv，再重跑本脚本生成证据。
"""
import csv
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import core  # noqa: E402

GOLD_PATH = os.path.join(HERE, "资料", "金标源", "语法金标样本.csv")
OUT_PATH = os.path.join(HERE, "data", "语法规则证据.json")
POOL_PATH = os.path.join(HERE, "资料", "金标源", "语法金标候选池.csv")

# 判定“命中”用全量触发（正则 + LTP 候选触发全部算真实命中点，与文档“抽查真实命中点”一致）
HIT_FIELD = "grammar_triggered"

# 人工标注价值最高的规则类别（词类规则按词性命中，无需抽样验收）
POOL_PROJECTS = ("句子的类型", "特殊句型", "固定格式", "口语格式", "固定短语")


def build_pool(docx_dir, out_path=POOL_PATH, max_per_no=20, max_sentences=4000):
    """从语料抽取官方可信【句式类】规则的命中句，输出待标注模板（金标列留空）。"""
    import docx_io
    _, _, grammar_rules = core._get_gf0025()
    targets = sorted({r["no"] for r in grammar_rules
                      if r.get("no") and r.get("trusted")
                      and r.get("project") in POOL_PROJECTS},
                     key=lambda x: int(x) if x.isdigit() else 10**9)
    docx_files = sorted((f for f in os.listdir(docx_dir)
                         if f.lower().endswith(".docx") and not docx_io.is_generated_report_file(f)),
                        key=docx_io.natural_key)
    pool = {}
    sentence_count = 0
    for fname in docx_files:
        path = os.path.join(docx_dir, fname)
        for _num, _name, body, _has_ex in docx_io.iter_lessons_from_docx(path, fname):
            for sent in core.split_sentences(body):
                if sentence_count >= max_sentences or len(pool) >= len(targets) * max_per_no:
                    break
                sentence_count += 1
                res = core.analyze_foreign_ltp(sent)
                for e in res.get("grammar_events") or []:
                    no = e.get("no")
                    if no in targets and no not in pool:
                        pool[no] = []
                    if no in targets and len(pool.get(no, [])) < max_per_no:
                        pool[no].append({
                            "fname": fname, "no": no, "level": e.get("level"),
                            "name": e.get("name"), "sentence": sent,
                            "detector_type": e.get("detector_type"),
                        })
            if sentence_count >= max_sentences or len(pool) >= len(targets) * max_per_no:
                break
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["句子", "语法点编号", "金标"])
        for no in targets:
            for row in pool.get(no, []):
                w.writerow([row["sentence"], no, ""])  # 金标留空待人工标注
    print(f"目标编号 {len(targets)} 个；扫描 {sentence_count} 句；样本 {sum(len(v) for v in pool.values())} 条 → {out_path}")
    for no in targets:
        print(f"  {no}: {len(pool.get(no, []))} 条")


def main():
    if "--build-pool" in sys.argv:
        import argparse
        ap = argparse.ArgumentParser(description="生成金标候选样本池")
        ap.add_argument("--build-pool", action="store_true")
        ap.add_argument("--dir", default=os.path.join(HERE, "资料"), help="语料 docx 目录")
        args = ap.parse_args()
        build_pool(args.dir)
        return
    rows = []
    with open(GOLD_PATH, encoding="utf-8-sig") as f:
        reader = list(csv.reader(f))[1:]
    for r in reader:
        if len(r) < 3 or not r[1].strip():
            continue
        rows.append((r[0].strip(), r[1].strip(), int(r[2].strip())))
    evidence = {}
    fns = 0
    for text, no, gold in rows:
        res = core.analyze_foreign_ltp(text)
        triggered = {n for lv in res[HIT_FIELD] for n in lv}
        hit = no in triggered
        ev = evidence.setdefault(no, {"samples": 0, "errors": 0, "tp": 0, "fp": 0, "fn": 0})
        ev["samples"] += 1
        if gold == 1 and hit:
            ev["tp"] += 1
        elif gold == 0 and hit:
            ev["fp"] += 1
            ev["errors"] += 1
        elif gold == 1 and not hit:
            ev["fn"] += 1
            fns += 1
        print(f"{no} 金标={gold} 命中={hit} | {text[:36]}")
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(evidence, f, ensure_ascii=False, indent=2)
    print("=" * 60)
    print(f"金标样本 {len(rows)} 条（漏报 {fns} 条，不计入精确率）→ {OUT_PATH}")
    for no, ev in sorted(evidence.items(), key=lambda kv: int(kv[0])):
        ok = (ev["samples"] >= core.GRAMMAR_TRUSTED_MIN_SAMPLE
              and ev["errors"] <= core.GRAMMAR_TRUSTED_MAX_ERRORS)
        print(f"  {no}: samples={ev['samples']} errors={ev['errors']} → "
              f"{'可升级可信' if ok else '样本不足/错误过多 → 保持候选'}")


if __name__ == "__main__":
    main()
