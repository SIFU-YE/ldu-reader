#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""LTP 语法检测器正/反例验证。"""
import importlib.util, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, '.libs312'))

from transformers import PreTrainedTokenizerBase, PreTrainedTokenizer, BertTokenizer
def _batch_encode_plus(self,*a,**k): return self(*a,**k)
for C in (PreTrainedTokenizerBase, PreTrainedTokenizer, BertTokenizer):
    if not hasattr(C,'batch_encode_plus'): C.batch_encode_plus=_batch_encode_plus
from ltp import LTP

spec = importlib.util.spec_from_file_location('gld', os.path.join(HERE, 'grammar_ltp_detectors.py'))
gld = importlib.util.module_from_spec(spec)
spec.loader.exec_module(gld)

CASES = {
 '136': {
   'pos': ['谁都可以参加。', '他哪儿都想去。'],
   'neg': ['他为什么不来？', '谁来了？'],
 },
 '140': {
   'pos': ['孩子们个个都很开心。'],
   'neg': ['每个孩子都很开心。'],
 },
 '160': {
   'pos': ['哈哈，太好了。'],
   'neg': ['他哈哈大笑。'],
 },
 '162': {
   'pos': ['孩子们一个一个走进教室。'],
   'neg': ['他有一个苹果。'],
 },
 '172': {
   'pos': ['他参加比赛是好事。'],
   'neg': ['他有一本书。'],
 },
 '173': {
   'pos': ['我知道他来了。'],
   'neg': ['他读书。'],
 },
 '175': {
   'pos': ['他关上了门。'],
   'neg': ['他关门。'],
 },
 '176': {
   'pos': ['他跑出去了。'],
   'neg': ['他跑。'],
 },
 '181': {
   'pos': ['毕业以后，他当了老师。'],
   'neg': ['他以前来过。'],
 },
 '183': {
   'pos': ['他把门关上了。'],
   'neg': ['他把书。'],
 },
 '184': {
   'pos': ['门被他关上了。'],
   'neg': ['他给朋友写信。'],
 },
 '185': {
   'pos': ['我们一起去商店买东西。'],
   'neg': ['他读书。'],
 },
 '186': {
   'pos': ['妈妈让他去买东西。'],
   'neg': ['妈妈爱他。'],
 },
 '187': {
   'pos': ['这本书比那本书有意思。'],
   'neg': ['他比较高。'],
 },
 '190': {
   'pos': ['他一会儿看书，一会儿写字。'],
   'neg': ['他休息了一会儿。'],
 },
 '193': {
   'pos': ['首先，我们听写。然后，我们读课文。'],
   'neg': ['他首先想到了家。'],
 },
 '195': {
   'pos': ['他不仅聪明，而且努力。'],
   'neg': ['他不仅聪明。'],
 },
 '196': {
   'pos': ['他不是学生，就是老师。'],
   'neg': ['他不是学生。'],
 },
 '197': {
   'pos': ['你喜欢就是喜欢，不喜欢就是不喜欢。'],
   'neg': ['他跑得快，就是有点累。'],
 },
 '198': {
   'pos': ['要是下雨，我们就不去。'],
   'neg': ['要是下雨，我们不去。'],
 },
 '209': {
   'pos': ['你想去就去吧。'],
   'neg': ['他说完就走了。', '他想跳就跳。'],
 },
 '131': {
   'pos': ['孩子们都喜欢。'],
   'neg': ['他喜欢猫。'],
 },
 '154': {
   'pos': ['他为了学习来到中国。'],
   'neg': ['他学习中文。'],
 },
 '161': {
   'pos': ['他往学校走。'],
   'neg': ['红苹果很好吃。'],
 },
 '174': {
   'pos': ['他写的字很漂亮。'],
   'neg': ['红苹果。'],
 },
 '177': {
   'pos': ['他听得懂。'],
   'neg': ['他听。'],
 },
 '178': {
   'pos': ['他高兴得很。'],
   'neg': ['他高兴。'],
 },
 '179': {
   'pos': ['他去过三次。'],
   'neg': ['他去了。'],
 },
 '180': {
   'pos': ['他等了一天。'],
   'neg': ['他等了。'],
 },
 '182': {
   'pos': ['这个人我认识。'],
   'neg': ['她很大方。'],
 },
 '188': {
   'pos': ['他看书看累了。'],
   'neg': ['他看书。'],
 },
 '192': {
   'pos': ['他又高又帅。'],
   'neg': ['他又来了。'],
 },
 '199': {
   'pos': ['只有努力才能成功。'],
   'neg': ['只有他来了。'],
 },
 '200': {
   'pos': ['因为下雨，所以没去。'],
   'neg': ['他去了。'],
 },
 '201': {
   'pos': ['为了健康，他跑步。'],
   'neg': ['他跑步。'],
 },
 '202': {
   'pos': ['他下了课就回家。'],
   'neg': ['他下课了，然后回家。'],
 },
 '203': {
   'pos': ['他大约三十岁。'],
   'neg': ['他三十岁。'],
 },
 '204': {
   'pos': ['他一点儿也不累。'],
   'neg': ['他累了。'],
 },
 '205': {
   'pos': ['难道你不知道吗？'],
   'neg': ['你知道吗？'],
 },
 '210': {
   'pos': ['你笑什么笑！'],
   'neg': ['你笑什么？'],
 },
}

ltp = LTP(os.path.join(HERE, 'LTP/small'))
all_ok = True
for no, case in CASES.items():
    fn = gld.DETECTORS.get(no)
    if not fn:
        print(f'{no}: 无检测器'); all_ok=False; continue
    sents = case['pos'] + case['neg']
    out = ltp.pipeline(sents, tasks=['cws','pos','dep'])
    results=[]
    for k, s in enumerate(sents):
        tokens = gld.parse_sentence(out.cws[k], out.pos[k], out.dep[k])
        r = bool(fn(tokens))
        results.append(r)
    pos_ok = all(results[:len(case['pos'])])
    neg_ok = not any(results[len(case['pos']):])
    ok = pos_ok and neg_ok
    all_ok &= ok
    print(f"{no}: 正例{'PASS' if pos_ok else 'FAIL'} 反例{'PASS' if neg_ok else 'FAIL'} -> {'PASS' if ok else 'FAIL'}")
    if not ok:
        for s,r in zip(sents, results):
            print('   ', '触发' if r else '未触发', s)

print('总结果:', '全部通过' if all_ok else '存在未通过')
