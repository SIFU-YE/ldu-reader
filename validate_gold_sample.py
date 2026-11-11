#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从75篇语料中抽取的语法金标样本验证 LTP 检测器。"""
import csv, importlib.util, os, sys
HERE=os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0,HERE); sys.path.insert(0,os.path.join(HERE,'.libs312'))
from transformers import PreTrainedTokenizerBase, PreTrainedTokenizer, BertTokenizer
def _batch_encode_plus(self,*a,**k): return self(*a,**k)
for C in (PreTrainedTokenizerBase, PreTrainedTokenizer, BertTokenizer):
    if not hasattr(C,'batch_encode_plus'): C.batch_encode_plus=_batch_encode_plus
from ltp import LTP
spec=importlib.util.spec_from_file_location('gld', os.path.join(HERE,'grammar_ltp_detectors.py'))
gld=importlib.util.module_from_spec(spec); spec.loader.exec_module(gld)

ltp=LTP(os.path.join(HERE,'LTP','small'))
rows=list(csv.reader(open(os.path.join(HERE,'data','语法金标样本.csv'),encoding='utf-8-sig')))[1:]
correct=0; total=0
for sent,no,gold in rows:
    gold=int(gold)
    out=ltp.pipeline([sent],tasks=['cws','pos','dep'])
    tokens=gld.parse_sentence(out.cws[0],out.pos[0],out.dep[0])
    fn=gld.DETECTORS.get(no)
    pred=1 if fn and fn(tokens) else 0
    ok=pred==gold; total+=1; correct+=ok
    print(('PASS' if ok else 'FAIL'), no, '金标', gold, '预测', pred, sent[:50])
print(f'金标准确率: {correct}/{total} = {correct/total*100:.1f}%')
