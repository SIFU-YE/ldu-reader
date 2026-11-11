#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量 Word 报告 · v4 完全本地离线版（阅读难度评估）

支持两类输入：
  1. 普通整册/单个文件：3-2.docx
     - 每个 docx 单独生成一份分析报告
  2. 分章文件：3-2-1.docx、3-2-2.docx ...
     - 每个章节单独生成 x-x-x_分析报告.docx
     - 自动按前缀 x-x 合并生成 x-x.docx
     - 对合并后的整册文本生成 x-x_分析报告.docx

用法:
  双击运行                       # 处理 exe 同目录下所有 docx
  python batch_report_v4.py --dir ./资料
"""

import argparse
import csv
import hashlib
import json
import os
import re
import statistics
import sys
import time

# 优先使用仓库内 .libs 离线依赖（python -m pip --target .libs）
_LOCAL_LIBS = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".libs")
if os.path.isdir(_LOCAL_LIBS) and _LOCAL_LIBS not in sys.path:
    sys.path.insert(0, _LOCAL_LIBS)

import core
import report
from docx_io import app_dir, extract_docx_paragraphs, is_generated_report_file

# 输入约定：仅分章文件 x-x-x.docx（如 1-1-1.docx、3-2-10.docx）
CHAPTER_RE = re.compile(r"^(\d+-\d+)-(\d+)\.docx$", re.I)
# 无整册/普通文件分支：输入非 x-x-x.docx 的文件直接跳过，避免混入口径


def wps_word_count(text):
    """与 WPS 中文文档“字数”接近的统计口径：
    汉字 + 中文标点/全角字符（如“”、。，！？等），不计空格与英文/数字。"""
    return sum(1 for c in text
               if "一" <= c <= "鿿"
               or "　" <= c <= "〿"
               or "＀" <= c <= "￯")


def _safe_stem(filename):
    stem = os.path.splitext(filename)[0]
    return "".join(c for c in stem if c not in r'\/:*?"<>|').strip() or "docx"


def _natural_key(name):
    """自然排序：3-2-10.docx 排在 3-2-2.docx 之后。"""
    return [int(part) if part.isdigit() else part
            for part in re.split(r"(\d+)", name)]


def _classify_docx(fname):
    """返回 (前缀, 章节号)；非分章文件（x-x-x.docx 之外）返回 None，调用方跳过。"""
    m = CHAPTER_RE.match(fname)
    if m:
        return m.group(1), m.group(2)
    return None


def _make_merged_docx(path, title, chapter_items):
    """把多个章节文本合并成一个 Word 文档。
    chapter_items: [(fname, text), ...]
    """
    from docx import Document
    doc = Document()
    doc.add_heading(title, level=0)
    for fname, text in chapter_items:
        doc.add_heading(_safe_stem(fname), level=1)
        for para in text.split("\n"):
            if para.strip():
                doc.add_paragraph(para.strip())
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    doc.save(path)
    return path


def _append_summary(summary_rows, fname, ftype, wps_chars, score, level,
                    stats, seconds, status=""):
    """统一追加 CSV 行（当前输入仅分章文件：类型=分章/汇总）。"""
    summary_rows.append([
        fname,
        ftype,
        wps_chars if wps_chars != "" else "",
        score if score != "" else "",
        level if level != "" else "",
        stats.get("total_chars", "") if stats else "",
        stats.get("unique_chars", "") if stats else "",
        stats.get("total_words", "") if stats else "",
        stats.get("unique_words", "") if stats else "",
        stats.get("sentences", "") if stats else "",
        stats.get("avg_sent_len", "") if stats else "",
        seconds if seconds != "" else "",
        status,
    ])


# ---------------------------------------------------------------------------
# 逐课结果 / 整合统计 / 语料追溯（V35）
# ---------------------------------------------------------------------------

def _sha256_text(text):
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _union_by_level(payloads):
    """多课明细（每课为 7 级 list）→ {level: 排序去重列表} + 各级总数（跨课并集）。"""
    union = {}
    total = 0
    for lv in range(7):
        s = set()
        for p in payloads:
            s.update(p[lv] or [])
        union[str(lv + 1)] = sorted(s)
        total += len(s)
    return union, total


def integrated_stats(records):
    """整合统计（业务口径：每分章文件=一课；只统计分析成功的课）。

    records: [ {ok, fname, score, level, stats, char_detail, word_detail,
                 grammar_covered, ...} ]，失败记录 ok=False 不计入分值统计。
    返回：分值分布（均值/中位数/标准差/最小/最大/全部分值）、等级课数、
    全课字/词/语法并集、字词次数累加。
    """
    oks = [r for r in records if r.get("ok")]
    failed = [r.get("fname") for r in records if not r.get("ok")]
    out = {
        "lesson_count": len(records),
        "ok_count": len(oks),
        "failed": failed,
    }
    if not oks:
        out.update({"score": None, "level_counts": {},
                    "char_unique_union": None, "word_unique_union": None,
                    "grammar_covered_union": None, "totals": None})
        return out
    scores = [r["score"] for r in oks]
    level_counts = {}
    for r in oks:
        level_counts[r["level"]] = level_counts.get(r["level"], 0) + 1
    char_union, char_total = _union_by_level([list(r.get("char_detail") or []) for r in oks])
    word_union, word_total = _union_by_level([list(r.get("word_detail") or []) for r in oks])
    gram_union, gram_total = _union_by_level([list(r.get("grammar_covered") or []) for r in oks])
    totals = {}
    for key in ("total_chars", "total_words", "total_hanzi_words",
                "total_word_tokens", "sentences", "unique_chars", "unique_words"):
        totals[key] = sum((r.get("stats") or {}).get(key, 0) for r in oks)
    out.update({
        "score": {
            "mean": round(statistics.mean(scores), 4),
            "median": round(statistics.median(scores), 4),
            "stdev": round(statistics.pstdev(scores), 4) if len(scores) > 1 else 0.0,
            "min": round(min(scores), 4),
            "max": round(max(scores), 4),
            "all": [round(s, 4) for s in scores],
        },
        "level_counts": dict(sorted(level_counts.items(), key=lambda kv: kv[0])),
        "char_unique_union": char_union,
        "char_unique_union_total": char_total,
        "word_unique_union": word_union,
        "word_unique_union_total": word_total,
        "grammar_covered_union": gram_union,
        "grammar_covered_union_total": gram_total,
        "totals": totals,
    })
    return out


def _write_lesson_records(records, out_dir):
    """写出 逐课结果.json（每分章一课，含失败记录）与 整合统计.json。"""
    records_path = os.path.join(out_dir, "逐课结果.json")
    with open(records_path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    stats_path = os.path.join(out_dir, "整合统计.json")
    with open(stats_path, "w", encoding="utf-8") as f:
        json.dump(integrated_stats(records), f, ensure_ascii=False, indent=2)
    return records_path, stats_path


# ---------------------------------------------------------------------------
# 语法四件套 CSV（V36：逐条证据 + 规则状态表）
# ---------------------------------------------------------------------------

def _write_integrated_report(records, out_path, prefix):
    """V39：整册整合统计报告（每课结果表 + 整合统计 + 并集摘要）。

    依据《程序运行逻辑说明（修订版）》：整册结果 = 各课分值平均；
    不重算整册文本（不再“拼成单篇重新评级”）。
    """
    from docx import Document
    from docx.shared import Pt, Cm

    doc = Document()
    doc.add_heading(f"{prefix} 阅读难度整合报告", level=0)
    para = doc.add_paragraph()
    para.add_run("统计口径：每个分章文件 = 一课独立评估；整册整合 = 各课分值"
                 "均值/中位数/标准差 + 等级课数 + 全课字词语法并集；"
                 "不把整册拼成单篇重新评级。")
    para.runs[0].font.size = Pt(10)

    integ = integrated_stats(records)
    doc.add_heading("一、每课结果", level=1)
    t = doc.add_table(rows=1, cols=8)
    t.style = "Table Grid"
    t.rows[0].cells[0].text = "文件"
    for i, h in enumerate(["分值", "级别", "总字数", "字种数", "词种数", "句子数", "平均句长"]):
        t.rows[0].cells[i + 1].text = h
    for r in records:
        if not r.get("ok"):
            cells = t.add_row().cells
            cells[0].text = r.get("fname", "")
            cells[1].text = "分析失败"
            continue
        s = r.get("stats") or {}
        cells = t.add_row().cells
        cells[0].text = r.get("fname", "")
        cells[1].text = f"{r['score']:.4f}"
        cells[2].text = r.get("level", "")
        cells[3].text = str(s.get("total_chars", ""))
        cells[4].text = str(s.get("unique_chars", ""))
        cells[5].text = str(s.get("unique_words", ""))
        cells[6].text = str(s.get("sentences", ""))
        cells[7].text = str(s.get("avg_sent_len", ""))

    doc.add_heading("二、整册整合统计", level=1)
    sc = integ.get("score") or {}
    lc = integ.get("level_counts") or {}
    tot = integ.get("totals") or {}
    rows = [
        ("课数（成功/总）", f"{integ.get('ok_count')}/{integ.get('lesson_count')}"),
        ("分值均值", sc.get("mean")),
        ("分值中位数", sc.get("median")),
        ("分值标准差", sc.get("stdev")),
        ("分值最小", sc.get("min")),
        ("分值最大", sc.get("max")),
        ("各难度等级课数", "；".join(f"{k}×{v}" for k, v in lc.items()) or "—"),
        ("全课字种并集", integ.get("char_unique_union_total")),
        ("全课词种并集", integ.get("word_unique_union_total")),
        ("全课可信语法点并集", integ.get("grammar_covered_union_total")),
        ("总字数/总词数/总句数",
         f"{tot.get('total_chars')}/{tot.get('total_words')}/{tot.get('sentences')}"),
    ]
    t2 = doc.add_table(rows=0, cols=2)
    t2.style = "Table Grid"
    for k, v in rows:
        cells = t2.add_row().cells
        cells[0].text = str(k)
        cells[1].text = str(v)

    doc.add_heading("三、结果文件", level=1)
    for line in ("逐课结果.json、整合统计.json、语料清单.csv、语料段落.jsonl、"
                 "语法四件套 CSV（{prefix}_语法*）均在同一输出目录"):
        p = doc.add_paragraph(line)
        p.runs[0].font.size = Pt(10)
    doc.save(out_path)


def _write_grammar_csvs(records, out_dir, prefix):
    """按【逐课结果记录】聚合输出四份语法 CSV（V39：lesson 级聚合，含 lesson_id）。

    1) {prefix}_语法可信命中_逐条.csv：status=trusted 事件；
    2) {prefix}_语法候选触发_逐条.csv：status=candidate 事件；
    3) {prefix}_语法候选触发_按等级汇总.csv：候选编号按等级计数；
    4) {prefix}_语法规则状态表.csv：全量 572 条 规则/检测器/证据/状态。
    """
    import collections
    import core as _core
    import grammar_ltp_detectors as _gld
    events = []
    for r in records:
        for e in (r.get("grammar_events") or []):
            ev = dict(e)
            ev.setdefault("lesson_id", r.get("fname", ""))
            events.append(ev)
    base = os.path.join(out_dir, f"{prefix}_语法")
    with open(base + "可信命中_逐条.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["文件", "句号", "等级", "编号", "语法点", "检测类型", "匹配片段"])
        for e in events:
            if e.get("status") == "trusted":
                w.writerow([e.get("lesson_id"), e.get("sentence_id"), e.get("level"),
                            e.get("no"), e.get("name"), e.get("detector_type"),
                            (e.get("matched_span") or "")[:100]])
    with open(base + "候选触发_逐条.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["文件", "句号", "等级", "编号", "语法点", "检测类型", "匹配片段"])
        for e in events:
            if e.get("status") == "candidate":
                w.writerow([e.get("lesson_id"), e.get("sentence_id"), e.get("level"),
                            e.get("no"), e.get("name"), e.get("detector_type"),
                            (e.get("matched_span") or "")[:100]])
    with open(base + "候选触发_按等级汇总.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["等级", "编号", "语法点", "候选命中次数"])
        counter = collections.Counter()
        for e in events:
            if e.get("status") == "candidate":
                counter[(e.get("level"), e.get("no"), e.get("name"))] += 1
        for (lv, no, name), n in sorted(counter.items(), key=lambda kv: (str(kv[0][0]), str(kv[0][1]))):
            w.writerow([lv, no, name, n])

    # 规则状态表：全量 572 条（gpi），合并 正则规则/检测器/样本证据/教材候选池句数
    import os as _os
    _gpi = _core._grammar_point_index()
    _, _, _grm = _core._get_gf0025()
    rule_by_no = {r.get("no"): r for r in _grm if r.get("no")}
    detect_nos = set(_gld.DETECTORS) | set(_gld.LEVEL_DETECTORS)
    pool_counts = collections.Counter()
    _pool_path = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)),
                               "资料", "金标源", "语法金标候选池.csv")
    if _os.path.exists(_pool_path):
        for _row in csv.reader(open(_pool_path, encoding="utf-8-sig")):
            if len(_row) >= 2 and _row[1].strip().isdigit():
                pool_counts[_row[1].strip()] += 1
    with open(base + "规则状态表.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["编号", "等级", "语法点", "来源", "正则可信",
                    "LTP检测器", "LTP可信", "样本数", "错误数", "教材候选句数", "状态"])
        for no in sorted(_gpi, key=lambda x: int(x) if x.isdigit() else 10**9):
            info = _gpi[no]
            rule = rule_by_no.get(no)
            src = "official"
            if rule and rule.get("no") in _core.GRAMMAR_REGEX_SUPPLEMENT:
                src = "supplement"
            regex_trusted = bool(rule and rule.get("trusted"))
            has_det = no in detect_nos
            det_trusted = has_det and _gld.is_trusted(no)
            samples, errors = (_core._get_rule_evidence().get(no) or {}).get("samples", 0), \
                              (_core._get_rule_evidence().get(no) or {}).get("errors", 0)
            if not rule and not has_det:
                status = "暂不支持（无正则/无检测器）"
            elif regex_trusted:
                status = "正则可信" + ("（样本验收）" if samples else "（配置身份/未验收）")
            elif det_trusted:
                status = "仅LTP检测器可信（正则候选）"
            else:
                status = "候选"
            w.writerow([no, info["level"], info["name"], src,
                        "是" if regex_trusted else "否",
                        "有" if has_det else "无",
                        "可信" if det_trusted else "候选",
                        samples, errors, pool_counts.get(no, 0), status])


def main():
    ap = argparse.ArgumentParser(description="v4 完全本地离线 · 阅读难度评估（对外汉语，按 docx 全文分析）")
    ap.add_argument("--dir", default=None, help="docx 所在目录（默认 = 程序所在目录）")
    ap.add_argument("--out-dir", default="分析报告", help="报告输出目录，默认 ./分析报告")
    ap.add_argument("--no-pause", action="store_true", help="运行完不暂停")
    args = ap.parse_args()

    # 扫描目录优先级：
    #   1) --dir 显式指定（绝对化，相对路径按当前工作目录解析）
    #   2) 程序(exe/脚本)所在目录——docx 与程序放同一文件夹即可双击运行
    #   3) 兜底：当前工作目录（兼容从其他目录用命令行启动的场景）
    if args.dir:
        base = os.path.abspath(args.dir)
    else:
        base = app_dir()
        if not any(f.lower().endswith(".docx") for f in os.listdir(base)):
            cwd = os.getcwd()
            if os.path.abspath(cwd) != os.path.abspath(base) and \
                    any(f.lower().endswith(".docx") for f in os.listdir(cwd)):
                base = cwd
    docx_files = sorted((f for f in os.listdir(base)
                         if f.lower().endswith(".docx") and not is_generated_report_file(f)),
                        key=_natural_key)
    if not docx_files:
        print(f"未在 {base} 找到 .docx 文件", file=sys.stderr)
        print("提示：请将 .docx 文档放到本程序所在文件夹（与程序同目录），再运行本程序。",
              file=sys.stderr)
        if not args.no_pause:
            input("\n按回车键退出...")
        sys.exit(1)

    out_dir = os.path.join(base, args.out_dir)
    os.makedirs(out_dir, exist_ok=True)
    corpus_dir = os.path.join(out_dir, "语料")
    os.makedirs(corpus_dir, exist_ok=True)

    print(f"v4 完全本地离线模式（LTP 分词）| 统计范围: 每个分章文件（x-x-x.docx）全文整体")
    print(f"扫描: {base} | 发现 {len(docx_files)} 个 docx\n")

    merged_corpus_parts = []
    summary_rows = []
    chapter_groups = {}   # prefix -> [dict]
    ok_count = 0
    lesson_records = []   # 逐课结果（每分章一课，含失败）
    trace_rows = []       # 语料清单 CSV 行
    para_lines = []       # 语料段落 jsonl 行

    for fname in docx_files:
        path = os.path.join(base, fname)
        klass = _classify_docx(fname)
        if klass is None:
            print(f"  [跳过] {fname}: 非分章文件（输入约定为 x-x-x.docx）")
            _append_summary(summary_rows, fname, "跳过", "", "", "", None, "", "非分章文件，跳过")
            continue
        prefix, chapter_no = klass
        print(f"== {fname} [分章]")
        try:
            paras = extract_docx_paragraphs(path)
        except Exception as exc:
            print(f"  [跳过] {type(exc).__name__}: {exc}")
            _append_summary(summary_rows, fname, "分章", "", "", "", None, "", "解析失败")
            lesson_records.append({"fname": fname, "prefix": prefix,
                                   "chapter_no": chapter_no, "ok": False,
                                   "error": f"{type(exc).__name__}: {exc}"})
            continue

        # 语料追溯：段落级 jsonl + 文件级清单行（哈希/段数/字数）
        for pi, para in enumerate(paras):
            para_lines.append({"file": fname, "prefix": prefix,
                               "chapter_no": chapter_no,
                               "para_index": pi, "text": para})
        full_text = "\n".join(paras).strip()
        if not full_text:
            print("  [跳过] 未提取到文本")
            _append_summary(summary_rows, fname, "分章", "", "", "", None, "", "未提取到文本")
            lesson_records.append({"fname": fname, "prefix": prefix,
                                   "chapter_no": chapter_no, "ok": False,
                                   "error": "未提取到文本"})
            continue

        wps_chars = wps_word_count(full_text)
        print(f"  全文 WPS 字数: {wps_chars}；开始 LTP 分析...")
        merged_corpus_parts.append(full_text)
        trace_rows.append([fname, prefix, chapter_no, _sha256_text(full_text),
                           len(paras), len(full_text), wps_chars])

        # 保存该 docx 的完整语料（全文整体，不删除任何段落）
        safe_name = _safe_stem(fname)
        with open(os.path.join(corpus_dir, f"{safe_name}.全文.txt"), "w", encoding="utf-8") as f:
            f.write(full_text)

        # 分章文本先进入合并组，即使单章 LTP 分析失败，也不影响合并文档/汇总报告
        chapter_entry = {
            "fname": fname,
            "text": full_text,
            "result": None,
            "stats": {},
            "wps_chars": wps_chars,
            "seconds": "",
        }
        chapter_groups.setdefault(prefix, []).append(chapter_entry)

        t0 = time.time()
        try:
            result = core.analyze_foreign_ltp(full_text)
        except Exception as exc:
            print(f"  [分析失败] {type(exc).__name__}: {exc}")
            _append_summary(summary_rows, fname, "分章", wps_chars, "", "", None, "", "分析失败")
            lesson_records.append({"fname": fname, "prefix": prefix,
                                   "chapter_no": chapter_no, "ok": False,
                                   "error": f"{type(exc).__name__}: {exc}",
                                   "sha256": _sha256_text(full_text),
                                   "total_chars": len(full_text)})
            continue

        seconds = round(time.time() - t0, 1)
        ok_count += 1

        # 按 docx 输出 Word 报告
        report_path = os.path.join(out_dir, f"{safe_name}_分析报告.docx")
        report.make_report(result, report_path,
                           doc_title=f"{fname} 阅读难度分析报告（全文整体）")

        s = result.get("stats", {})
        chapter_entry["result"] = result
        chapter_entry["stats"] = s
        chapter_entry["seconds"] = seconds
        lesson_records.append({
            "fname": fname, "prefix": prefix, "chapter_no": chapter_no,
            "ok": True, "sha256": _sha256_text(full_text),
            "wps_chars": wps_chars, "seconds": seconds,
            "score": result["score"], "level": result["level"],
            "stats": s,
            "char_detail": [list(x) for x in result.get("char_detail") or []],
            "word_detail": [list(x) for x in result.get("word_detail") or []],
            "grammar_covered": [list(x) for x in result.get("grammar_covered") or []],
            "grammar_events": list(result.get("grammar_events") or []),  # V39：并入逐课证据（含 lesson=fname）
        })
        _append_summary(summary_rows, fname, "分章", wps_chars,
                        f"{result['score']:.4f}", result["level"], s, seconds)
        print(f"  分值: {result['score']:.4f} | 级别: {result['level']} | "
              f"总字数 {s.get('total_chars','?')} | 字种 {s.get('unique_chars','?')} | "
              f"总词数 {s.get('total_words','?')} | 词种 {s.get('unique_words','?')} | "
              f"句子 {s.get('sentences','?')} | 平均句长 {s.get('avg_sent_len','?')} | "
              f"耗时 {seconds}s")
        print(f"  报告: {report_path}")

    # 生成分章汇总：合并章节文本为 x-x.docx（语料存档），并生成【整合统计报告】
    # V39（第1项方案A）：按《程序运行逻辑说明（修订版）》——整册结果=各课分值平均；
    # 不再把整册拼成单篇重新评级（原 merged_text → analyze_foreign_ltp 路径已删除）。
    for prefix in sorted(chapter_groups.keys(), key=_natural_key):
        parts = sorted(chapter_groups[prefix], key=lambda x: _natural_key(x["fname"]))
        if not parts:
            continue
        merged_text = "\n\n".join(p["text"] for p in parts).strip()
        if not merged_text:
            continue

        # 1) 合并后的原始文档（仅语料存档，不参与评级）
        merged_docx = os.path.join(out_dir, f"{prefix}.docx")
        _make_merged_docx(merged_docx, f"{prefix} 合并文档",
                          [(p["fname"], p["text"]) for p in parts])
        with open(os.path.join(corpus_dir, f"{prefix}.全文.txt"), "w", encoding="utf-8") as f:
            f.write(merged_text)
        print(f"== 合并 {prefix}: {len(parts)} 个章节 -> {merged_docx}（语料存档）")

        # 2) 整合统计报告（逐课结果 + 平均/并集；不重算整册文本）
        recs = [r for r in lesson_records if r.get("prefix") == prefix]
        summary_report = os.path.join(out_dir, f"{prefix}_分析报告.docx")
        _write_integrated_report(recs, summary_report, prefix)
        _write_grammar_csvs(recs, out_dir, prefix)
        integ = integrated_stats(recs)
        sc = integ.get("score") or {}
        tot = integ.get("totals") or {}
        wps_total = sum((r.get("wps_chars") or 0) for r in recs if r.get("ok"))
        _append_summary(summary_rows, os.path.basename(merged_docx), "整合",
                        wps_total,
                        f"{sc['mean']:.4f}" if sc.get("mean") is not None else "",
                        core.gf0025_level(sc["mean"]) if sc.get("mean") is not None else "",
                        {"total_chars": tot.get("total_chars", ""),
                         "unique_chars": tot.get("unique_chars", ""),
                         "total_words": tot.get("total_words", ""),
                         "unique_words": tot.get("unique_words", ""),
                         "sentences": tot.get("sentences", ""),
                         "avg_sent_len": round(tot.get("total_chars", 0) / tot.get("sentences", 1), 1)
                         if tot.get("sentences") else ""},
                        "")
        print(f"  整合报告: {summary_report}（分值平均 {sc.get('mean')}，"
              f"共 {integ.get('ok_count')} 课）")

    if ok_count == 0:
        print("错误: 所有分析均失败（含分章汇总）", file=sys.stderr)
        sys.exit(1)

    # 保存合并全量语料（仅用于存档，不参与分析）
    full_merged = "\n\n".join(merged_corpus_parts).strip()
    with open(os.path.join(out_dir, "合并语料_全文.txt"), "w", encoding="utf-8") as f:
        f.write(full_merged)

    # 输出汇总 CSV（按 docx / 汇总）
    summary_path = os.path.join(out_dir, "分析结果汇总.csv")
    with open(summary_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["文件名", "类型", "全文WPS字数", "难度分值", "级别",
                    "总字数(非空格)", "字种数", "总词数", "词种数",
                    "句子数", "平均句长", "耗时(s)", "备注"])
        w.writerows(summary_rows)

    # 逐课结果 + 整合统计（V35：每分章=一课）
    records_path, stats_path = _write_lesson_records(lesson_records, out_dir)

    # 语料追溯（V35）：语料清单.csv + 语料段落.jsonl
    trace_path = os.path.join(out_dir, "语料清单.csv")
    with open(trace_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["文件名", "前缀", "章节号", "sha256", "段落数", "文本字符数", "WPS字数"])
        w.writerows(trace_rows)
    paras_path = os.path.join(out_dir, "语料段落.jsonl")
    with open(paras_path, "w", encoding="utf-8") as f:
        for line in para_lines:
            f.write(json.dumps(line, ensure_ascii=False) + "\n")

    print(f"逐课结果: {records_path}")
    print(f"整合统计: {stats_path}")
    print(f"语料清单: {trace_path}")
    print(f"语料段落: {paras_path}")

    print("\n" + "=" * 56)
    print(f"完成: {ok_count} 次分析成功")
    print(f"汇总 CSV: {summary_path}")
    print(f"语料目录: {corpus_dir}")
    print(f"合并语料: {os.path.join(out_dir, '合并语料_全文.txt')}")
    print("=" * 56)
    print("注: 每个 x-x-x.docx 单独生成 x-x-x_分析报告.docx；"
          "同组章节合并为 x-x.docx（仅语料存档），并生成 x-x_分析报告.docx = "
          "《整合统计报告》（每课结果 + 分值平均/等级课数/全课并集，不重算整册文本）。")
    if not args.no_pause:
        input("\n按回车键退出...")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        # 顶层兜底：双击运行时任何未捕获异常都要显示错误并暂停，
        # 防止控制台窗口瞬间关闭（用户感知为"闪退"）。
        import traceback
        traceback.print_exc()
        print("\n程序发生错误，请将本窗口内容截图反馈。", file=sys.stderr)
        try:
            input("\n按回车键退出...")
        except Exception:
            pass
        sys.exit(1)
