#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
词汇全等级覆盖与复现（3-1~3-5 全部75篇，自动分词口径，不含人工校对）
标准词表：《国际中文教育中文水平等级标准》词汇表（GF0025_词汇_backup.csv 条目口径）
分词策略：LTP 分词 + 词表OOV切分（与三级词汇复现同一策略）
输出：分析报告/词汇等级覆盖复现_全等级.csv

口径说明：本脚本是【词表口径】验证（标准词表覆盖/词形数），使用 LTP 独立复现；
【生产分析流程】的分词/词性/分值也以 core.analyze_foreign_ltp（LTP）输出为准，
两者口径一致（均 LTP）。
"""
import argparse, csv, os, re, sys
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, '.libs'))

import docx_io as batch_ldu
import core

FILES = ['3-1', '3-2', '3-3', '3-4', '3-5']
OUT = os.path.join(HERE, '分析报告', '词汇等级覆盖复现_全等级.csv')


def clean(w):
    return re.split(r'[（(]', w)[0].strip()


def seg_ltp_plain(text):
    """分词：统一使用 LTP（全项目只 LTP）。"""
    res = []
    chunks = core._chunk_text(text)
    ltp_results = core._ltp_analyze_chunks(chunks)
    if ltp_results is None:
        raise RuntimeError("LTP 不可用，无法进行词表覆盖分析。")
    for r in ltp_results:
        for w in r.get("cws", []):
            w = w.strip()
            if core._hanzi_of(w):
                res.append(w)
    return res


def main():
    ap = argparse.ArgumentParser(description='词汇全等级覆盖与复现（结巴分词，任意 docx 目录）')
    ap.add_argument('--dir', default=os.path.join(HERE, '资料'), help='docx 目录')
    ap.add_argument('--out', default=OUT, help='输出 CSV')
    args = ap.parse_args()
    base = os.path.abspath(args.dir)
    docx_files = sorted((f for f in os.listdir(base)
                         if f.lower().endswith('.docx') and not batch_ldu.is_generated_report_file(f)),
                        key=batch_ldu.natural_key)
    if not docx_files:
        print(f'未在 {base} 找到 .docx')
        sys.exit(1)
    all_lessons = []
    for fn in docx_files:
        if batch_ldu.has_volume_chapter_files(base, fn):
            continue
        path = os.path.join(base, fn)
        for _num, _name, body, _has_ex in batch_ldu.iter_lessons_from_docx(path, fn):
            if body:
                all_lessons.append(body)
    merged = "\n".join(all_lessons).strip()

    cl, wl, gr = core._get_gf0025()
    freq = Counter()
    for w in seg_ltp_plain(merged):
        freq[w] += 1

    rows = []
    print('词汇等级覆盖与复现（普通分词×官方唯一词表，不反向切分不重复计条目）')
    print('等级 | 标准词形 | 覆盖词形 | 覆盖率 | 出现1 | 出现2 | 出现3 | 出现4 | 出现≥5 | 复现≥2')
    for lv in range(1, 8):
        lv_set = {w for w, l in wl.items() if l == lv}
        covered_set = lv_set & set(freq)
        covered_n = len(covered_set)
        occ = [0] * 5
        for w in covered_set:
            n = freq[w]
            idx = min(n, 5) - 1
            occ[idx] += 1
        rep = sum(occ[1:])
        rows.append([lv, len(lv_set), covered_n,
                     f'{covered_n / len(lv_set) * 100:.2f}%',
                     occ[0], occ[1], occ[2], occ[3], occ[4], rep])
        name = core.LEVEL_NAMES[lv-1]
        print(f'{name:<6} | {len(lv_set):>6} | {covered_n:>6} | '
              f'{covered_n / len(lv_set) * 100:>6.2f}% | '
              f'{occ[0]:>5} | {occ[1]:>5} | {occ[2]:>5} | {occ[3]:>5} | {occ[4]:>6} | {rep:>6}')

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['等级', '标准词形', '覆盖词形', '覆盖率', '出现1次', '出现2次',
                    '出现3次', '出现4次', '出现≥5次', '复现≥2次'])
        w.writerows(rows)
    print('已写入:', args.out)

if __name__ == "__main__":
    main()
