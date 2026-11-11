#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""LTP 整篇文本桥接：stdin JSON -> stdout JSON。供 Python3.11 主流程调用 py3.12 环境。"""
import json, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '.libs312'))

from transformers import PreTrainedTokenizerBase, PreTrainedTokenizer, BertTokenizer
def _batch_encode_plus(self, *args, **kwargs):
    return self(*args, **kwargs)
for C in (PreTrainedTokenizerBase, PreTrainedTokenizer, BertTokenizer):
    if not hasattr(C, 'batch_encode_plus'):
        C.batch_encode_plus = _batch_encode_plus

from ltp import LTP


def main():
    data = json.load(sys.stdin)
    texts = data.get('texts', [])
    if not texts:
        print(json.dumps({'results': []}))
        return
    texts = [t for t in texts if t and t.strip()]
    ltp = LTP(os.path.join(HERE, 'LTP', 'small'))
    results = []
    BATCH = 16  # 每批16块，避免整批过大导致超时/内存过高
    for i in range(0, len(texts), BATCH):
        batch = texts[i:i + BATCH]
        out = ltp.pipeline(batch, tasks=['cws', 'pos', 'dep'])
        for j in range(len(batch)):
            results.append({
                'cws': out.cws[j],
                'pos': out.pos[j],
                'dep': {
                    'head': [int(x) for x in out.dep[j]['head']],
                    'label': out.dep[j]['label'],
                },
            })
    print(json.dumps({'results': results}, ensure_ascii=False))


if __name__ == '__main__':
    main()
