#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""语法规则抽样验收：对每个可信规则抽取最多10个命中点，输出 CSV 供人工复核。
论文方法：每条规则抽查10个命中点，错误≤1个才接受。本脚本执行抽样部分；判断由人工完成。"""
import argparse, csv, os, re, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE); sys.path.insert(0, os.path.join(HERE, '.libs'))
import docx_io as batch_ldu, core

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--dir', default=os.path.join(HERE, '资料'))
    ap.add_argument('--level', type=int, default=None, help='仅抽查该级')
    ap.add_argument('--out', default=os.path.join(HERE, '分析报告', '语法规则抽样复核.csv'))
    args = ap.parse_args()
    base = os.path.abspath(args.dir)
    docx_files = sorted((f for f in os.listdir(base)
                         if f.lower().endswith('.docx') and not batch_ldu.is_generated_report_file(f)),
                        key=batch_ldu.natural_key)
    texts = []
    for f in docx_files:
        if batch_ldu.has_volume_chapter_files(base, f):
            continue
        path = os.path.join(base, f)
        for _num, _name, body, _has_ex in batch_ldu.iter_lessons_from_docx(path, f):
            texts.append(body)
    merged = '\n'.join(texts)
    char_levels, word_levels, grammar_rules = core._get_gf0025()
    rows = []
    for rule in grammar_rules:
        if not rule.get('trusted') or not rule.get('no'):
            continue
        if args.level and rule['level'] != args.level:
            continue
        matches = []
        for text in texts:
            for sent in core.split_sentences(text):
                m = re.search(rule['regex'], sent)
                if m:
                    matches.append(sent.strip())
                if len(matches) >= 10:
                    break
            if len(matches) >= 10:
                break
        for i, sent in enumerate(matches, 1):
            rows.append([rule['no'], rule['level'], rule['name'], rule['regex'][:60], i, sent[:200]])
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f); w.writerow(['编号', '级别', '规则名', '正则', '样本序号', '命中句']); w.writerows(rows)
    print('已写入', args.out, '共', len(rows), '条样本')

if __name__ == '__main__':
    main()
