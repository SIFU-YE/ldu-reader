#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用 LTP 重新分词，复现论文词汇表3和三级词覆盖/复现（不含人工校对）。"""
import argparse, csv, datetime, hashlib, json, os, re, statistics, subprocess, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, '.libs312'))

# transformers 5 兼容：补 batch_encode_plus
from transformers import PreTrainedTokenizerBase, PreTrainedTokenizer, BertTokenizer
def _batch_encode_plus(self, *args, **kwargs):
    return self(*args, **kwargs)
for C in (PreTrainedTokenizerBase, PreTrainedTokenizer, BertTokenizer):
    if not hasattr(C, 'batch_encode_plus'):
        C.batch_encode_plus = _batch_encode_plus

from ltp import LTP
import docx_io as batch_ldu

DEFAULT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '资料')


def clean(w):
    return re.split(r'[（(]', w)[0].strip()


def load_official_unique():
    d = {}
    with open(os.path.join(HERE, 'GF0025_词汇_official.csv'), encoding='utf-8') as f:
        next(f)
        for row in csv.reader(f):
            if len(row) < 3:
                continue
            lv_text = row[1].strip()
            w = row[2].strip()
            lv = None
            for prefix, n in (("一", 1), ("二", 2), ("三", 3), ("四", 4), ("五", 5), ("六", 6)):
                if lv_text.startswith(prefix):
                    lv = n
                    break
            if lv is None and lv_text.startswith("高等"):
                lv = 7
            if lv and w:
                d[w] = lv
    return d


def main():
    ap = argparse.ArgumentParser(description='LTP 词汇覆盖/复现（全部 docx 合并为一篇整体分析）')
    ap.add_argument('--dir', default=DEFAULT_DIR, help='docx 目录，默认 ./资料')
    ap.add_argument('--out', default=os.path.join(HERE, '分析报告', 'ltp_vocabulary_result.json'), help='输出 JSON')
    args = ap.parse_args()
    base = os.path.abspath(args.dir)
    docx_files = sorted((f for f in os.listdir(base)
                         if f.lower().endswith('.docx') and not batch_ldu.is_generated_report_file(f)),
                        key=batch_ldu.natural_key)
    if not docx_files:
        print(f'未在 {base} 找到 .docx')
        sys.exit(1)

    # 合并所有 docx 的所有课文正文为一篇（不切块）
    merged_parts = []
    lesson_count = 0
    for fn in docx_files:
        if batch_ldu.has_volume_chapter_files(base, fn):
            continue
        path = os.path.join(base, fn)
        for _num, _name, body, _has_ex in batch_ldu.iter_lessons_from_docx(path, fn):
            if body:
                merged_parts.append(body)
                lesson_count += 1
    merged = "\n".join(merged_parts).strip()
    if not merged:
        print('未提取到课文正文')
        sys.exit(1)

    import core as _core
    wl = _core._get_gf0025()[1]  # 标准词表（GF0025_词汇.csv 11129条展开后）
    chunks = _core._sentence_list(merged)
    results = _core._ltp_analyze_chunks(chunks)
    if results is None:
        print('LTP 不可用：请设置 LTP_PYTHON 或运行 download_ltp_model.py 下载模型')
        sys.exit(1)
    all_word_freq = {}
    detail = [set() for _ in range(7)]
    for res in results:
        for w in res.get('cws', []):
            w = w.strip()
            if not re.fullmatch(r'[\u4e00-\u9fff]+', w):
                continue
            lv = wl.get(w)
            if lv:
                detail[lv - 1].add(w)
            all_word_freq[w] = all_word_freq.get(w, 0) + 1

    levels = []
    for lv in range(1, 8):
        lv_set = {w for w, l in wl.items() if l == lv}
        covered = lv_set & set(all_word_freq)
        occ = [0] * 5
        for w in covered:
            n = all_word_freq[w]
            idx = min(n, 5) - 1
            occ[idx] += 1
        levels.append({
            'level': lv, 'standard': len(lv_set), 'covered': len(covered),
            'coverage_pct': round(len(covered) / len(lv_set) * 100, 2),
            'occ1': occ[0], 'occ2': occ[1], 'occ3': occ[2], 'occ4': occ[3], 'occ_ge5': occ[4],
            'reproduced_ge2': sum(occ[1:]),
        })
        print(lv, levels[-1])

    def _git_head():
        try:
            return subprocess.check_output(['git', '-C', HERE, 'rev-parse', 'HEAD'],
                                           text=True, stderr=subprocess.DEVNULL).strip()
        except Exception:
            return ''
    # 词表代次指纹：缓存校验用（reproduce_paper.check_ltp_cache 会与当前词表对比，
    # 词表变更后旧缓存的 covered/standard 虽自洽但不再是当前口径，需提示重跑）
    level_counts = [sum(1 for w, l in wl.items() if l == lv) for lv in range(1, 8)]
    result = {
        'dir': base, 'merged': True, 'lessons': lesson_count, 'merged_chars': len(merged),
        'levels': levels,
        'generated_at': datetime.datetime.now().isoformat(),
        'git_commit': _git_head(),
        'model': 'LTP/small',
        'corpus_sha256': hashlib.sha256(merged.encode('utf-8')).hexdigest(),
        'chunk_count': len(chunks),
        'word_level_fingerprint': {'counts': level_counts, 'total': len(wl)},
    }
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print('saved', args.out)

if __name__ == "__main__":
    main()
