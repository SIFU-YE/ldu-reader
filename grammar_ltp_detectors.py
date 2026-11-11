#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基于 LTP 句法特征的语法点检测器（阶段3）

用法:
  LTP_PYTHON=/path/to/python3.12 python grammar_ltp_detectors.py --dir ./资料 --out 分析报告/语法_LTP检测.csv

说明:
  - 使用 LTP 分词/词性/依存句法，对《等级标准》一~七九级语法点做“特征触发”检测。
  - 与正则检测不同，检测依据是 LTP 的 SBV/VOB/CMP/DBL/COO/POB/ADV 等依存关系。
  - 三级部分检测器计入可信覆盖；其他等级统一作候选提示（报告单独列出）。
  - 可对任意 docx 目录运行；无 LTP 时不可用（会提示）。
"""
import argparse, csv, datetime, hashlib, json, os, re, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
# 本模块只做“检测器”逻辑，不直接 import LTP；LTP 调用统一走 core._ltp_analyze_chunks。
# 因此不要把 .libs312 插入 sys.path，否则在 Python3.11 主流程里会污染后续 LTP 导入。

import docx_io as batch_ldu

# 结果补语常用词
RESULT_CMP = {'到', '住', '走', '完', '好', '懂', '见', '开', '上', '下', '出', '起来'}
TREND_CMP = {'出来', '出去', '过来', '过去', '回来', '回去', '进来', '进去', '起来', '上来', '上去', '下来', '下去'}
DEGREE_CMP = {'很', '多', '极了', '死了', '不得了', '慌', '厉害', '坏', '透'}
TIME_QUANT = {'年', '天', '小时', '分钟', '个月', '星期', '周', '月', '日'}
ACTION_QUANT = {'次', '遍', '场', '下', '回', '趟'}
DBL_VERBS = {'叫', '派', '请', '让', '使', '令'}
COMPARE_PREP = {'比', '跟', '和', '与', '同'}


def parse_sentence(words, pos, dep):
    """返回 {word:(pos, head_idx, rel)} 以及列表。"""
    tokens = []
    for i, (w, p) in enumerate(zip(words, pos)):
        tokens.append({'i': i, 'w': w, 'pos': p,
                       'head': int(dep['head'][i]) if dep else -1,
                       'rel': dep['label'][i] if dep else ''})
    return tokens


def child_of(tokens, idx, rel=None):
    return [t for t in tokens if t['head'] == idx + 1]  # LTP head 是 1-based

# ---------------------------------------------------------------------------
# 各语法点检测器
# ---------------------------------------------------------------------------
def det_132(tokens):
    """语素后缀：-儿、-家、-们、-头、-子。

    修复 P17：
    - 独立后缀 token（LTP 标 k/u，如“们/子/头”被切出）触发；
    - 词尾为后缀字且整词非固化名词（LTP 不标 n/nh 等）触发；
    - 排除“孩子/儿子/桌子/房子”等固化名词（LTP 标 n）。
    """
    for t in tokens:
        w = t['w']
        p = t['pos']
        # 独立后缀 token（长度1也有效，如“们/k”“子/u”“头/u”）
        if len(w) == 1 and w in ('儿', '家', '们', '头', '子') and p.startswith(('k', 'h', 'u')):
            return True
        if len(w) < 2:
            continue
        if not w.endswith(('儿', '家', '们', '头', '子')):
            continue
        # 固化名词词尾（孩子/儿子/桌子/房子…）不视为构词后缀
        if p in ('n', 'nh', 'nm', 'ns', 'nt', 'nz', 'nx'):
            continue
        # 词性为后缀类或非名词的构词（如 科学家→n? LTP 可能标 n，但仍属构词后缀）
        # 保守候选：词性为 k/h/u，或末字后缀且词性非名词/代词/副词
        if p.startswith(('k', 'h', 'u')):
            return True
        if p not in ('r', 'd'):
            return True
    return False


def det_141(tokens):
    """量词重叠：AA / 一X一X（颗颗、件件、一棵棵、一张张…）。

    修复 P17：仅当重叠模式出现时触发；LTP 对“一棵棵”常标 m（数词），
    对“条条/件件”常标 q（量词），故词性接受 q/m/数词，但必须满足重叠词形
    （AA 或 一+X+X），排除“妈妈/谢谢”等非量词叠词。
    """
    for t in tokens:
        w = t['w']
        p = t['pos']
        q_like = p.startswith(('q', 'm', 'qg')) or p in ('q', 'm')
        # 双字叠词：AA（条条/件件/次次…）且为量词/数词类
        if len(w) == 2 and w[0] == w[1] and q_like:
            return True
        # 三字“一X一Y”式（一棵棵/一张张/一个个）：LTP 可能整体切一词
        if len(w) == 3 and w[0] == '一' and w[1] == w[2] and q_like:
            return True
        # 相邻 token：一 + X + X（LTP 切成一/m + 棵/q + 棵/q 或 一/m + 棵/m）
        if w in ('一', '壹') and q_like:
            idx = t['i']
            nxt1 = tokens[idx + 1] if idx + 1 < len(tokens) else None
            nxt2 = tokens[idx + 2] if idx + 2 < len(tokens) else None
            if (nxt1 and nxt2 and nxt1['w'] == nxt2['w']
                    and nxt1['pos'].startswith(('q', 'm')) and nxt2['pos'].startswith(('q', 'm'))):
                return True
    return False


def det_155(tokens):
    """介词引出目的原因：因为/由于/为了"""
    for t in tokens:
        if t['w'] in ('因为', '由于', '为了'):
            return True
    return False


def det_161(tokens):
    for t in tokens:
        if t['w'] == '哈哈':
            return True
    return False


def det_162(tokens):
    """结构类型：介宾/兼语/同位短语（用 LTP POB/DBL 关系近似，不再用ATT）"""
    rels = {t['rel'] for t in tokens}
    return 'POB' in rels or 'DBL' in rels


def det_173(tokens):
    """动词/形容词短语作主语：主语节点本身带有 SBV 或 VOB（从句作主语）"""
    for t in tokens:
        if t['rel'] == 'SBV':
            children = child_of(tokens, t['i'])
            for c in children:
                if c['rel'] in ('SBV', 'VOB', 'CMP'):
                    return True
    return False


def det_174(tokens):
    """动词/形容词短语、主谓短语作宾语：宾语节点本身带有 SBV 或 CMP"""
    for t in tokens:
        if t['rel'] == 'VOB':
            children = child_of(tokens, t['i'])
            for c in children:
                if c['rel'] in ('SBV', 'CMP'):
                    return True
    return False


def det_175(tokens):
    """动词性/主谓短语作定语：ATT 节点本身带有 SBV/VOB/CMP"""
    for t in tokens:
        if t['rel'] == 'ATT':
            children = child_of(tokens, t['i'])
            for c in children:
                if c['rel'] in ('SBV', 'VOB', 'CMP'):
                    return True
    return False


def det_176(tokens):
    """结果补语：CMP 且补语为结果词"""
    for t in tokens:
        if t['rel'] == 'CMP' and t['w'] in RESULT_CMP:
            head = tokens[t['head'] - 1] if t['head'] > 0 else None
            if head and head['pos'].startswith('v'):
                return True
    return False


def det_177(tokens):
    """趋向补语：CMP 且补语为趋向词"""
    for t in tokens:
        if t['rel'] == 'CMP' and t['w'] in TREND_CMP:
            return True
    return False


def det_178(tokens):
    """可能补语：V + 得/不 + V/A（LTP 中常见 RAD/CMP）"""
    for i, t in enumerate(tokens):
        if t['rel'] == 'RAD' and t['w'] in ('得', '不'):
            if i + 1 < len(tokens) and tokens[i + 1]['pos'] in ('v', 'a'):
                return True
    for t in tokens:
        if t['rel'] == 'CMP' and t['w'] in ('了', '到', '起', '下', '上', '完', '住', '动', '走', '回', '来', '去'):
            head = tokens[t['head'] - 1] if t['head'] > 0 else None
            if head and head['pos'].startswith('v'):
                # 可能补语常见：V得/V不 + CMP
                rads = [x for x in tokens if x['head'] == head['i'] + 1 and x['rel'] == 'RAD' and x['w'] in ('得', '不')]
                if rads:
                    return True
    return False


def det_179(tokens):
    """程度补语：CMP 且补语为程度词"""
    for t in tokens:
        if t['rel'] == 'CMP' and t['w'] in DEGREE_CMP:
            return True
    return False


def det_180(tokens):
    """数量补语：CMP 且补语为数量词/数词"""
    for t in tokens:
        if t['rel'] == 'CMP' and (t['pos'] == 'q' or re.fullmatch(r'[0-9]+(次|遍|场|下|回|趟)', t['w'])):
            return True
    return False


def det_181(tokens):
    """时量补语：CMP 且补语含时间量词/量词q前接数词"""
    for i, t in enumerate(tokens):
        if t['rel'] == 'CMP':
            if re.search(r'[0-9]+(年|天|小时|分钟|个月|星期|周|月|日)|一会儿|很久|半天', t['w']):
                return True
            if t['pos'] == 'q' and i > 0 and tokens[i - 1]['pos'] == 'm':
                return True
    return False


def det_182(tokens):
    """数量补语（间隔）：CMP + 以后/之后/以前/之前/来/以来"""
    for t in tokens:
        if t['rel'] == 'CMP' and re.search(r'(以后|之后|以前|之前|以来)$', t['w']):
            return True
    return False


def det_183(tokens):
    """主谓谓语句：LTP 前置宾语(FOB)与主谓关系(SBV)共现"""
    rels = {t['rel'] for t in tokens}
    return 'FOB' in rels and 'SBV' in rels


def det_184(tokens):
    """把字句：把 + POB + 动词（POB 后必须有动词性成分）"""
    for i, t in enumerate(tokens):
        if t['w'] == '把' and (t['pos'] == 'p' or t['rel'] == 'ADV'):
            pob = [c for c in child_of(tokens, t['i']) if c['rel'] == 'POB']
            for c in pob:
                for j in range(c['i'] + 1, len(tokens)):
                    if tokens[j]['pos'].startswith('v'):
                        return True
    return False


def det_185(tokens):
    """被动句：被/叫/让 + POB + 动词；或“被”为ADV且直连动词（无施事）"""
    for i, t in enumerate(tokens):
        if t['w'] in ('被', '叫', '让') and (t['pos'] == 'p' or t['rel'] == 'ADV'):
            pob = [c for c in child_of(tokens, t['i']) if c['rel'] == 'POB']
            for c in pob:
                for j in range(c['i'] + 1, len(tokens)):
                    if tokens[j]['pos'].startswith('v'):
                        return True
            # 无施事被动：被/叫/让 直接支配动词
            if t['head'] > 0:
                head = tokens[t['head'] - 1]
                if head['pos'].startswith('v'):
                    return True
    return False


def det_186(tokens):
    """连动句：两个动词性词语构成 COO/后续动作"""
    coo = [t for t in tokens if t['rel'] == 'COO']
    if coo:
        for t in coo:
            head = tokens[t['head'] - 1] if t['head'] > 0 else None
            if head and head['pos'].startswith('v') and t['pos'].startswith('v'):
                return True
    return False


def det_187(tokens):
    """兼语句：使令动词 + DBL + 动词性成分（第二动词与使令动词同支配）。

    V41 收紧：原先 head 允许【任意动词】（如"教/打电话找我/催"被误判为兼语1
    表使令），误触发率高；现仅接受使令动词 DBL_VERBS={叫,派,请,让,使,令}——
    "女儿还教我在网上买东西""一天，有个朋友打电话找我"等不再触发（186 主编号
    与 251/328 候选同时受益）。
    """
    for t in tokens:
        if t['rel'] == 'DBL' and t['head'] > 0:
            head = tokens[t['head'] - 1]
            if head['w'] in DBL_VERBS:
                for other in tokens:
                    if other['i'] != t['i'] and other['head'] == t['head'] and other['pos'].startswith('v'):
                        return True
    return False


def det_188(tokens):
    """比较句：比/跟/和/与/同 + POB"""
    for t in tokens:
        if t['w'] in COMPARE_PREP and t['pos'] == 'p':
            if any(c['rel'] == 'POB' for c in child_of(tokens, t['i'])):
                return True
    return False


def det_189(tokens):
    """重动句：V + O + V + CMP（宾语不要求为动词）"""
    for i, t in enumerate(tokens):
        if t['rel'] == 'VOB':
            head = tokens[t['head'] - 1] if t['head'] > 0 else None
            if head and head['pos'].startswith('v'):
                for j in range(i + 1, len(tokens)):
                    if tokens[j]['w'] == head['w'] and tokens[j]['rel'] in ('HED', 'VOB', 'CMP', 'COO'):
                        return True
    return False


def det_191(tokens):
    return any(t['w'] == '一会儿' for t in tokens)


def det_193(tokens):
    """又……又……：至少两个“又”"""
    return sum(1 for t in tokens if t['w'] == '又') >= 2


def det_200(tokens):
    """只有……才……"""
    ws = {t['w'] for t in tokens}
    return '只有' in ws and '才' in ws


def det_201(tokens):
    return any(t['w'] in ('所以', '因此', '由于') for t in tokens)


def det_202(tokens):
    return any(t['w'] == '为了' for t in tokens)


def det_203(tokens):
    """紧缩复句：……了……（就）……：同句“了”在“就”前，且二者之间无标点"""
    ws = [t['w'] for t in tokens]
    pos = [t['pos'] for t in tokens]
    if '了' not in ws or '就' not in ws:
        return False
    # 取首个“了”与【最后一个】“就”：紧缩句“……了……就……”的“就”常在
    # 句末（如“他一吃完饭就去了”），ws.index('就') 只会取到首个“就”，
    # “他就吃了饭马上就去”这类首“就”在“了”前的句子会被漏检。
    i = ws.index('了')
    k = len(ws) - 1 - ws[::-1].index('就')
    if i >= k:
        return False
    return not any(pos[m] == 'wp' for m in range(i + 1, k))


def det_204(tokens):
    return any(t['w'] in ('大概', '大约', '左右', '前后') for t in tokens)


def det_205(tokens):
    """用“一点儿也不……”表示强调：一点儿 + 也 + 不/没"""
    ws = [t['w'] for t in tokens]
    for i in range(len(ws) - 2):
        if ws[i] == '一点儿' and ws[i + 1] == '也' and ws[i + 2] in ('不', '没'):
            return True
    return False


def det_206(tokens):
    return any(t['w'] == '难道' for t in tokens)


def det_211(tokens):
    """X什么（啊）：X什么X格式"""
    ws = [t['w'] for t in tokens]
    for i in range(len(ws) - 2):
        if ws[i + 1] == '什么' and ws[i] == ws[i + 2]:
            return True
    return False


def det_137(tokens):
    """疑问代词任指：疑问代词 + 都/也"""
    q_words = {'谁', '什么', '哪儿', '哪里', '怎么', '怎样', '多少', '几'}
    has_q = any(t['w'] in q_words for t in tokens)
    has_dy = any(t['w'] in ('都', '也') for t in tokens)
    return has_q and has_dy


def det_163(tokens):
    """数量重叠：数词+量词+数词+量词"""
    pat = re.compile(r'^[0-9一二三四五六七八九十]+(个|只|条|张|本|支|块|件|双|岁|年|天|次|遍|口|位|名|种|些|点)$')
    for i in range(len(tokens) - 1):
        if pat.match(tokens[i]['w']) and pat.match(tokens[i + 1]['w']):
            return True
    return False


def det_182b(tokens):
    """数量补语（间隔）：动词/时间 + 以后/之后/以前/之前/以来"""
    for i, t in enumerate(tokens):
        if t['w'] in ('以后', '之后', '之前', '以来') and i > 0:
            prev = tokens[i - 1]
            if prev['pos'].startswith('v') or prev['pos'] in ('nt', 'nd', 't', 'm'):
                return True
    return False


def det_191b(tokens):
    return sum(1 for t in tokens if t['w'] == '一会儿') >= 2


def det_194b(tokens):
    ws = {t['w'] for t in tokens}
    return '首先' in ws and '然后' in ws


def det_196b(tokens):
    ws = {t['w'] for t in tokens}
    return ('不仅' in ws or '不光' in ws) and ('还' in ws or '而且' in ws)


def det_197b(tokens):
    ws = [t['w'] for t in tokens]
    for i in range(len(ws) - 1):
        if ws[i] == '不' and ws[i + 1] == '是':
            for j in range(i + 2, len(ws) - 1):
                if ws[j] == '就' and ws[j + 1] == '是':
                    return True
    return False


def det_198b(tokens):
    for i, t in enumerate(tokens):
        if t['w'] == '就是' and 0 < i < len(tokens) - 1:
            if tokens[i - 1]['w'] == tokens[i + 1]['w']:
                return True
    return False


def det_199b(tokens):
    ws = {t['w'] for t in tokens}
    return '要是' in ws and '就' in ws


def det_210b(tokens):
    """X就X（点儿）吧：要求句中有“吧”，且“就”前后同形单字动词/形容词"""
    if not any(t['w'] == '吧' for t in tokens):
        return False
    for i, t in enumerate(tokens):
        if t['w'] == '就' and 0 < i < len(tokens) - 1:
            a, b = tokens[i - 1], tokens[i + 1]
            if a['w'] == b['w'] and len(a['w']) == 1 and a['pos'] in ('v', 'a'):
                return True
    return False


# 以下检测器特征过宽，只作候选提示，不计入“可信LTP覆盖”
WEAK_DETECTOR_NOS = {
    # 结构性弱/依赖句内局部关系的检测器
    '161', '182', '185',
    # 后缀/AA/比较/重动等词面规则过宽，只作候选
    '131', '140', '187', '188',
    # 整篇送入LTP后，以下模式不再具备句内边界，只作候选提示
    '136', '190', '193', '195', '196', '197', '198', '199', '200', '201', '202', '203', '204', '205', '209', '210',
    # 词面级候选
    '192',
}


DETECTORS = {
    '131': det_132, '136': det_137, '140': det_141, '154': det_155, '160': det_161, '161': det_162,
    '162': det_163, '172': det_173, '173': det_174, '174': det_175, '175': det_176, '176': det_177,
    '177': det_178, '178': det_179, '179': det_180, '180': det_181, '181': det_182b,
    '182': det_183, '183': det_184, '184': det_185, '185': det_186, '186': det_187,
    '187': det_188, '188': det_189, '190': det_191b, '192': det_193, '193': det_194b, '195': det_196b, '196': det_197b,
    '197': det_198b, '198': det_199b, '199': det_200, '209': det_210b,
    '200': det_201, '201': det_202, '202': det_203, '203': det_204, '204': det_205,
    '205': det_206, '210': det_211,
}


def _has(tokens, w):
    return any(t['w'] == w for t in tokens)


def det_you(tokens):
    """“有”字句：存在/领有"""
    for t in tokens:
        if t['w'] == '有' and t['pos'].startswith('v'):
            rels = {c['rel'] for c in tokens if c['head'] == t['i'] + 1}
            if 'SBV' in rels or 'VOB' in rels:
                return True
    return False


def det_cunxian(tokens):
    """存现句：处所词 + 有/着 + 名词"""
    rels = {t['rel'] for t in tokens}
    return ('POB' in rels or 'ADV' in rels) and (_has(tokens, '有') or _has(tokens, '着'))


def det_double_obj(tokens):
    """双宾语句：同一动词支配两个 VOB"""
    vob_heads = [t['head'] for t in tokens if t['rel'] == 'VOB']
    return len(vob_heads) >= 2 and len(set(vob_heads)) < len(vob_heads)


def det_shi_de(tokens):
    """是……的句"""
    return _has(tokens, '是') and _has(tokens, '的')


def det_zhe(tokens):
    """持续态：动态助词“着”"""
    return any(t['w'] == '着' and t['pos'] == 'u' for t in tokens)


def det_guo(tokens):
    """经历态：动态助词“过”"""
    return any(t['w'] == '过' and t['pos'] == 'u' for t in tokens)


def det_yiwen(tokens):
    """疑问句
    V55 修正 '几' 误报：当句末无问号/吗/呢时，'几' LTP 标 m (数词) 应排除特指疑问判定；
    '什么/哪/哪儿/哪里' LTP 标 r (代词) 才算疑问；句末标点是更可靠的判别依据。
    V55 修正 '呢' 陈述句误判："呢"在无问号、无疑问代词时多为舒缓/继续语气
    （"他明天来呢。"），仅"吗"是强疑问标记；"呢"须配合问号或疑问代词才判疑问。
    """
    # 第一类：明确的疑问标记（句末问号 / 强疑问"吗"）→ 整句是疑问
    if any(t['w'] in ('？', '?') for t in tokens):
        return True
    if any(t['w'] == '吗' for t in tokens):
        return True
    # 第二类：句中含疑问代词（r 词性的代词 + 谁/什么/哪/哪儿/哪里/哪些/怎么）→ 特指疑问
    # （"呢"单独不构成疑问标记，由疑问代词或问号负责）
    _interrog_pronouns = {'谁', '什么', '哪', '哪儿', '哪里', '哪些', '怎么', '怎样', '为什么'}
    for t in tokens:
        if t['w'] in _interrog_pronouns and t['pos'] == 'r':
            return True
    # 第三类：含"多少"——LTP 通常标 m，但搭配"多少"整体作疑问
    if any(t['w'] == '多少' for t in tokens):
        return True
    return False


def det_qishi(tokens):
    """祈使句：无 SBV 主语，且含动词"""
    rels = {t['rel'] for t in tokens}
    if 'SBV' in rels:
        return False
    return any(t['pos'].startswith('v') for t in tokens)


def det_gantan(tokens):
    """感叹句"""
    return any(t['w'] in ('！', '!') for t in tokens)


def det_bijiao1(tokens):
    """比较句1：比/没有/有 + POB"""
    for t in tokens:
        if t['w'] in ('比', '没有', '有') and t['pos'] in ('p', 'v'):
            if any(c['rel'] == 'POB' for c in tokens if c['head'] == t['i'] + 1):
                return True
    return False


def det_jinxing(tokens):
    """进行态：在/正在/正 + 动词"""
    ws = [t['w'] for t in tokens]
    pos = [t['pos'] for t in tokens]
    for i in range(len(ws) - 1):
        if ws[i] in ('在', '正在', '正') and pos[i + 1].startswith('v'):
            return True
    return False


def det_zhengfan(tokens):
    """正反疑问：A不A/A没A"""
    ws = [t['w'] for t in tokens]
    for i in range(len(ws) - 2):
        if ws[i + 1] in ('不', '没') and ws[i] == ws[i + 2]:
            return True
    return False


def _has_buzhi(tokens):
    """识别“不知”的不同分词结果：不知 / 不+知 / 不知道……"""
    for t in tokens:
        if t['w'].startswith('不知'):
            return True
    ws = [t['w'] for t in tokens]
    for i in range(len(ws) - 1):
        if ws[i] == '不' and ws[i + 1].startswith('知'):
            return True
    return False


def det_402(tokens):
    """……便……：副词/连词“便”"""
    return any(t['w'] == '便' and t['pos'] in ('d', 'c') for t in tokens)


def det_407(tokens):
    """虽（然）……，但/可/却/也……"""
    ws = [t['w'] for t in tokens]
    return any(w in ('虽', '虽然') for w in ws) and any(w in ('但', '但是', '可', '可是', '却', '也') for w in ws)


def det_452(tokens):
    """不知……好"""
    ws = [t['w'] for t in tokens]
    return _has_buzhi(tokens) and any(w == '好' for w in ws)


def det_454(tokens):
    """无非/只不过/只是……而已/罢了"""
    ws = [t['w'] for t in tokens]
    return any(w in ('无非', '只不过', '只是') for w in ws) and any(w in ('而已', '罢了') for w in ws)


def det_455(tokens):
    """以……为……：'以'作介词并带 POB 宾语，且句中有动词'为'"""
    for t in tokens:
        if t['w'] == '以' and t['pos'] == 'p':
            if any(c['rel'] == 'POB' for c in child_of(tokens, t['i'])):
                if any(x['w'] == '为' and x['pos'].startswith('v') for x in tokens):
                    return True
    return False


def det_506(tokens):
    """不知X的"""
    return _has_buzhi(tokens) and any(t['w'] == '的' for t in tokens)


# 其他等级检测器：按标准编号映射（可信/候选由 LEVEL_DETECTOR_TRUSTED_NOS 决定）
LEVEL_DETECTORS = {
    # 一级
    '33': det_yiwen, '34': det_qishi, '35': det_gantan, '38': det_bijiao1,
    '42': det_jinxing, '48': det_zhengfan,
    # 二级
    '103': det_you, '105': det_186, '109': det_double_obj, '118': det_zhe,
    '119': det_guo, '108': det_shi_de,
    # 四级
    '248': det_184, '249': det_185, '251': det_187, '252': det_shi_de,
    # 五级
    '325': det_184, '326': det_185, '328': det_187, '329': det_188,
    # 六级
    '397': det_184, '398': det_185,
    '401': det_402, '406': det_407,
    # 七-九级
    '451': det_452, '453': det_454, '454': det_455, '505': det_506,
}

# 非三级检测器的可信集合：只有表层结构明确、且不易与其它等级混淆的检测器才计可信。
# 其余非三级检测器仍作候选提示，进入报告的“LTP候选命中”，但不进入可信覆盖与难度计分。
# 注意：下列编号也需与 data/GF0025_语法_with_regex.csv 对齐（曾因编号平移整体错位 1）。
# “不知X的”（505）的检测器仅按“有‘不知/不知道’分词且句中有‘的’”判定，
# 会把“他们不知道，这也是一个很累的工作”等普通否定句误判为可信高等语法 → 降为候选。
# 本集合与 core.GRAMMAR_UNRELIABLE_NOS（正则路径）【独立判定】：同一编号两条路径可
# 各有不同的可信结论（如 454“以……为……”正则过宽不可信、LTP 依存检测器更精确可信），
# 属有意设计；报告中“正则可信命中”与“LTP可信命中”分表展示，互不覆盖。
# 【V26 修正】326（被动句3：意念被动句，无“被/叫/让”显式标记）与 397（把字句4：表致使）
# 的语义无法由通用检测器 det_184/det_185（仅识别显式“把/被/叫/让/给”结构）证实——
# 检测器命中它们必然把普通把/被字句误计为可信命中；降为候选（仍触发，供人工复核）。
# 【V30 修正】把字句/被动句/兼语句各细分结构共用同一“大类检测器”
# （det_184 把、det_185 被动、det_187 兼语），只能证明“大类结构存在”、
# 无法区分细分语义（处置1/2/3/致使4、被动1/2/意念/4、使令/爱憎称谓/致使）。
# 若多个细分编号同时计为可信，同一句会在三/四/五/六级【重复记级】且抬高难度分值。
# 故仅保留【主编号】可信（把183、被动184、兼语186 = 各细分的起始/基础结构），
# 其余细分编号（248/249/251/325/328/398 等）降为候选提示，供人工复核。
LEVEL_DETECTOR_TRUSTED_NOS = {
    # 一级
    # V55：'33' 疑问句从可信集合移除 —— 用户评审确认 33 的逐句证据不准确
    # （"几"曾误报，修复后仍属语义级判定），且无金标证据 → 必须人工复核。
    # 保留 35/38/42/48（表层结构明确）。33 仍触发为候选，供人工复核。
    '35', '38', '42', '48',
    # 二级
    '103', '109', '118', '119',
    # 七-九级
    '454',
}


def is_trusted(no):
    """该编号的 LTP 检测器是否计入可信覆盖。"""
    if no in LEVEL_DETECTORS:
        return no in LEVEL_DETECTOR_TRUSTED_NOS
    return no not in WEAK_DETECTOR_NOS


def detect_triggered(tokens):
    """对单句 LTP 分析结果运行全部句法检测器，返回命中的标准语法点编号集合。

    tokens 由 parse_sentence(words, pos, dep) 生成。任何检测器异常只跳过该检测器，
    不影响主流程。"""
    triggered = set()
    for no, fn in DETECTORS.items():
        try:
            if fn(tokens):
                triggered.add(no)
        except Exception:
            continue
    for no, fn in LEVEL_DETECTORS.items():
        try:
            if fn(tokens):
                triggered.add(no)
        except Exception:
            continue
    return triggered


def detect_texts(texts, ltp):
    """整篇文本作为整体送入 LTP，不做分句/分段处理。"""
    triggered = {no: 0 for no in DETECTORS}
    for no in LEVEL_DETECTORS:
        triggered.setdefault(no, 0)
    for text in texts:
        if not text.strip():
            continue
        out = ltp.pipeline([text], tasks=['cws', 'pos', 'dep'])
        words = out.cws[0]
        pos = out.pos[0]
        dep = out.dep[0]
        tokens = parse_sentence(words, pos, dep)
        for no, fn in list(DETECTORS.items()) + list(LEVEL_DETECTORS.items()):
            try:
                if fn(tokens):
                    triggered[no] += 1
            except Exception:
                pass
    return triggered


def main():
    ap = argparse.ArgumentParser(description='LTP 语法特征检测器')
    ap.add_argument('--dir', default=os.path.join(HERE, '资料'), help='docx 目录')
    ap.add_argument('--out', default=os.path.join(HERE, '分析报告', '语法_LTP检测.csv'), help='输出CSV')
    args = ap.parse_args()
    base = os.path.abspath(args.dir)
    files = sorted((f for f in os.listdir(base)
                    if f.lower().endswith('.docx') and not batch_ldu.is_generated_report_file(f)),
                   key=batch_ldu.natural_key)
    if not files:
        print('未找到 docx')
        sys.exit(1)
    texts = []
    for f in files:
        if batch_ldu.has_volume_chapter_files(base, f):
            print('skip volume with chapters', f)
            continue
        try:
            path = os.path.join(base, f)
            for _num, _name, body, _has_ex in batch_ldu.iter_lessons_from_docx(path, f):
                texts.append(body)  # 统一不计标题
        except Exception as e:
            print('skip', f, e)
    merged = "\n".join(texts).strip()
    import core as _core
    chunks = _core._sentence_list(merged)
    print('检测语料:', len(texts), '篇，合并为', len(merged), '字，切为', len(chunks), '块')
    results = _core._ltp_analyze_chunks(chunks)
    if results is None:
        print('LTP 不可用：请设置 LTP_PYTHON 或运行 download_ltp_model.py 下载模型')
        sys.exit(1)
    trig = {no: 0 for no in DETECTORS}
    for no in LEVEL_DETECTORS:
        trig.setdefault(no, 0)
    for res in results:
        tokens = parse_sentence(res.get('cws', []), res.get('pos', []), res.get('dep', {}))
        for no, fn in list(DETECTORS.items()) + list(LEVEL_DETECTORS.items()):
            try:
                if fn(tokens):
                    trig[no] += 1
            except Exception:
                pass
    rows = [['编号', '命中块数', '是否计入可信LTP覆盖']]
    for no in sorted(trig, key=lambda x: int(x)):
        trusted = '是' if is_trusted(no) else '否'
        rows.append([no, trig[no], trusted])
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f); w.writerows(rows)
    print('已写入', args.out)
    meta_path = args.out.replace('.csv', '_meta.json')
    try:
        head = subprocess.check_output(['git', '-C', HERE, 'rev-parse', 'HEAD'],
                                       text=True, stderr=subprocess.DEVNULL).strip()
    except Exception:
        head = ''
    meta = {
        'generated_at': datetime.datetime.now().isoformat(),
        'git_commit': head,
        'model': 'LTP/small',
        'corpus_sha256': hashlib.sha256(merged.encode('utf-8')).hexdigest(),
        'chunk_count': len(chunks),
        'sentence_count': len(chunks),
    }
    with open(meta_path, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    print('已写入', meta_path)
    all_nos = set(DETECTORS) | set(LEVEL_DETECTORS)
    trusted_nos = {no for no in all_nos if is_trusted(no)}
    print('可信命中:', sum(1 for no in trusted_nos if trig.get(no, 0) > 0), '/', len(trusted_nos), '个可信检测器')
    for no in sorted(trig, key=lambda x: int(x)):
        if trig[no] == 0:
            print('  未触发', no)


if __name__ == '__main__':
    main()
