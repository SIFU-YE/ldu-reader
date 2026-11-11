#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地阅读难度分析核心（基于 python-readability-cn 源码复刻，无 1 万字限制）

- foreign（对外汉语）：程勇 GF0025 公式  难度=(1/3)*Σ n*(r_字+r_词+r_语法)
- native（母语汉语）：程勇母语可读性公式  38.36-45.65*字频/1e6+54.92*连词比-8.96*名词比
  +11.13*语义丰富度-12.34*动词比+0.012*句长方差+20*连词比
- 完整输出分级字/词/语法点各级别明细

依赖：标准库 + LTP（分词/词性统一使用 LTP，项目不再依赖 jieba）。
数据：python-readability-cn-main/readability_cn/data/
"""

import csv
import json
import os
import re
import shutil
import sys
from collections import Counter, defaultdict

# 优先使用仓库内 .libs 离线依赖（python -m pip --target .libs）；
# 同时兼容 install_ltp_env.py 安装的 .libs312（旧安装脚本目标目录）。
_HERE_DIR = os.path.dirname(os.path.abspath(__file__))
for _libs_dir in (".libs", ".libs312"):
    _libs_path = os.path.join(_HERE_DIR, _libs_dir)
    if os.path.isdir(_libs_path) and _libs_path not in sys.path:
        sys.path.insert(0, _libs_path)

if getattr(sys, "frozen", False):
    # PyInstaller 打包：GF0025 数据随 --add-data 打进 _MEIPASS/data
    DATA_DIR = os.path.join(sys._MEIPASS, "data")
else:
    # P0-5：优先仓库自带 data/（克隆后可运行），缺失时回退外部数据目录
    _local_data = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
    if os.path.isdir(_local_data):
        DATA_DIR = _local_data
    else:
        DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "python-readability-cn-main", "readability_cn", "data")

LEVEL_NAMES = ["一级", "二级", "三级", "四级", "五级", "六级", "七-九级"]
LEVEL_PREFIX = [("一", 1), ("二", 2), ("三", 3), ("四", 4), ("五", 5), ("六", 6)]


def _parse_level(text):
    """级别文字 -> 1..7（高等=7），未知返回 0"""
    for prefix, n in LEVEL_PREFIX:
        if text.startswith(prefix):
            return n
    return 7 if text.startswith("高等") else 0


# ---------------------------------------------------------------------------
# 词性 -> 《标准》词性类别（用于 (词形, 词性) 联合定级）
# ---------------------------------------------------------------------------
# 一级类目用首字母前缀匹配：LTP 词性细类（如 vd/vn、nr/ns）归入《标准》大类
_WORD_POS_CAT_PREFIXES = [
    ("名", ("n",)),
    ("动", ("v",)),
    ("形", ("a", "b")),
    ("副", ("d",)),
    ("介", ("p",)),
    ("连", ("c",)),
    ("助", ("u",)),
    ("叹", ("e",)),
    ("数", ("m",)),
    ("量", ("q",)),
    ("代", ("r",)),
    ("拟声", ("o",)),
    ("前缀", ("h",)),
    ("后缀", ("k",)),
]

# 词表/分词器中出现的非标准词性符号 -> 《标准》大类（精确匹配置于前缀匹配之前，
# 防止 `pr` 被 p 前缀误映射为“介”、`name` 被 n 前缀误映射为“名”）。
_WORD_POS_CAT_EXACT = {
    "pr": ("代",),       # 代词类符号（如“有些/有一些”曾被标 Pr）
    "name": ("名",),     # 人名/专名符号（如“老王/小王”曾被标 Name），归入名类
}


def _pos_to_word_cats(pos):
    """LTP 词性 -> 《标准》词性大类列表（一个词性只归一个大类；未知返回 []）"""
    p = (pos or "").strip().lower()
    if not p:
        return []
    cats = []
    for cat, prefixes in _WORD_POS_CAT_PREFIXES:
        if any(p.startswith(pref) for pref in prefixes):
            cats.append(cat)
    # 精确特殊符号覆盖前缀匹配结果（如 pr 不应为“介”）
    return list(_WORD_POS_CAT_EXACT.get(p, cats))


_POS_LABELS = {
    "a": "形", "b": "形", "d": "副", "n": "名", "v": "动", "q": "量", "m": "数",
    "p": "介", "c": "连", "u": "助", "e": "叹", "r": "代", "o": "拟声",
    "h": "前缀", "k": "后缀", "w": "标点", "wp": "标点",
}


def _word_detail_label(word, pos):
    """用于分级词明细展示的“词/词性”标签，避免同形多词性词在多级重复时无法区分。"""
    p = (pos or "").strip().lower()
    if p in _POS_LABELS:
        return f"{word}/{_POS_LABELS[p]}"
    cats = _pos_to_word_cats(p)
    if cats:
        return f"{word}/{cats[0]}"
    return word if not p else f"{word}/{p}"


def _word_entry_key(word, pos):
    """标准词表条目键：优先归一到《标准》词性大类，未知词性用原始词性。"""
    cats = _pos_to_word_cats(pos)
    if cats:
        return (word, cats[0])
    return (word, (pos or "").strip().lower())


def _lookup_word_level_detail(word_levels, word_pos_levels, word, pos):
    """(词形, 词性) 联合定级详情。

    返回 (等级, 是否多级候选, 定级来源)：
      - source = "pos"   ：(词形, 标准词性大类) 精确命中标准表；
      - source = "empty" ：标准表中该词无词性标记，按空词性命中；
      - source = "form"  ：词性不匹配，仅按词形兜底到最低级（保守口径）；
      - source = None    ：未定级。
    """
    cats = _pos_to_word_cats(pos)
    levels = []
    source = None
    for cat in cats:
        vals = word_pos_levels.get((word, cat))
        if vals:
            levels.extend(vals)
            source = "pos"
    if not levels:
        vals = word_pos_levels.get((word, ""))
        if vals:
            levels.extend(vals)
            source = "empty"
    if not levels:
        all_levels = sorted({lv for (w, _cat), vals in word_pos_levels.items()
                             if w == word for lv in vals})
        if all_levels:
            return all_levels[0], len(all_levels) > 1, "form"
        return 0, False, None
    levels = sorted(set(levels))
    return levels[0], len(levels) > 1, source


def _lookup_word_level(word_levels, word_pos_levels, word, pos):
    """(词形, 词性) 联合定级：优先用词性命中词表中的具体大类；
    若该 (词形, 词性) 在词表中有多个等级，返回 (最低级, True) 标记为多级候选；
    词性缺失/未命中时回退到词形最低级（保守口径）。
    返回 (等级, 是否多级候选)。"""
    lv, amb, _source = _lookup_word_level_detail(word_levels, word_pos_levels, word, pos)
    return lv, amb


def _format_effective_detail(exact_set, fallback_words, all_exact_forms=None):
    """把精确命中 (词形,词性) 集合与独立词形兜底集合格式化为统一展示列表。

    - 精确命中：标签为“词/词性”；
    - 独立兜底词形：仅当该词形未在任何等级精确命中时加入。
    all_exact_forms：全局精确命中词形集合（跨等级去重）；None 时用本级。
    """
    all_exact = all_exact_forms if all_exact_forms is not None else {w for (w, _c) in exact_set}
    out = {_word_detail_label(w, cat) for (w, cat) in exact_set}
    for w in fallback_words:
        if w not in all_exact:
            out.add(w)
    return out


# 专名/数字组合/分词残片：超纲词清单中不应直接当作人工教学词表的类型。
_PROPER_NOUN_POS = {"nr", "ns", "nt", "nz", "nh", "ni", "j", "nz", "nrt"}
_NUMBER_POS_PREFIXES = ("m",)
_CHINESE_NUM_CHARS = "零〇一二两三四五六七八九十百千万亿点几多"
_NUMBER_SUFFIXES = (
    "第", "月", "日", "号", "年", "点", "分", "秒", "元", "块", "角", "毛", "岁",
    "倍", "次", "遍", "个", "只", "条", "张", "本", "件", "双", "位", "名", "种",
    "些", "天", "周", "星期", "半",
)


def _is_proper_noun_pos(pos):
    """是否为专名/缩略类词性（人名、地名、机构、其他专名、简称等）。"""
    p = (pos or "").strip().lower()
    return p in _PROPER_NOUN_POS or p.startswith(("nr", "ns", "nt", "nz", "nh", "ni"))


def _is_number_like_word(word, pos=None):
    """是否为数字组合/数量表达式（如 二十一、第三、三月、五号）。"""
    w = (word or "").strip()
    if not w:
        return False
    p = (pos or "").strip().lower()
    if p.startswith(_NUMBER_POS_PREFIXES):
        return True
    # 纯汉字数字串
    if re.fullmatch(f"[{_CHINESE_NUM_CHARS}]+", w):
        return True
    # 数字 + 常见数量/时间/序数后缀
    if re.fullmatch(f"[第]?[{_CHINESE_NUM_CHARS}]+(?:{'|'.join(_NUMBER_SUFFIXES)})?", w):
        # 必须至少含一个数字汉字，避免把“天/周”等普通名词全筛掉
        return any(ch in _CHINESE_NUM_CHARS for ch in w)
    return False


def _is_single_char_fragment(word, pos=None):
    """单字未登录：很可能是分词残片/不成词语素，需人工复核。"""
    return len((word or "").strip()) == 1


def _is_phrase_like_word(word, word_levels):
    """若一个词表外多字串可完全切分为两个及以上标准词，则判为短语/分词结果。

    例如“很快”=“很”+“快”，不应直接作为独立教学词条。
    """
    w = (word or "").strip()
    if not w or len(w) < 2 or not word_levels:
        return False
    n = len(w)
    max_len = min(12, n)
    memo = {}

    def can_seg(i):
        if i == n:
            return True
        if i in memo:
            return memo[i]
        for j in range(i + 1, min(i + max_len, n) + 1):
            if w[i:j] in word_levels and can_seg(j):
                memo[i] = True
                return True
        memo[i] = False
        return False
    return can_seg(0)


def _classify_beyond_word(word, pos=None, word_levels=None):
    """将词表外纯汉字词分为：
    proper(专名/缩略)、number(数字组合/数量)、fragment(单字残片)、
    phrase(短语/分词结果)、ordinary(普通超纲词)。
    """
    if _is_proper_noun_pos(pos):
        return "proper"
    if _is_number_like_word(word, pos):
        return "number"
    if _is_single_char_fragment(word, pos):
        return "fragment"
    if word_levels is not None and _is_phrase_like_word(word, word_levels):
        return "phrase"
    return "ordinary"


def _add_beyond_word(buckets, word, pos=None, word_levels=None):
    """按类型写入超纲词桶。buckets: {"ordinary":{}, "proper":{}, "number":{}, "fragment":{}, "phrase":{}}"""
    cat = _classify_beyond_word(word, pos, word_levels)
    buckets.setdefault(cat, {}).setdefault(word, pos)


# ---------------------------------------------------------------------------
# GF0025 数据加载
# ---------------------------------------------------------------------------
def _fix_format_regex(regex, content, det):
    """P1-2 四字格/格式规则修复：贪婪跨距 `X(.+)Y(.+)` → 单字紧邻 `X(.)Y(.)`。
    只对形如「东A西B」「一A一B」的格式规则生效（内容/细目为 1-2 字+ A/B 占位）。
    P1-3 被动句"为X所Y/被X所Y"四字结构：间隔限 1-4 字且不跨标点
    （防止"因为…所以…"被误判为"为…所…"）。"""
    for src in (content, det):
        s = (src or "").strip()
        if re.match(r"^\S{1,2}A\S{0,2}B\S{0,2}$", s) and "A" in s and "B" in s:
            return regex.replace("(.+)", "(.)")
    if "被.+所.+" in regex or "为.+所.+" in regex:
        regex = (regex.replace("被.+所.+", "被[^，。]{1,4}所[^，。]")
                      .replace("为.+所.+", "为[^，。]{1,4}所[^，。]"))
    # A-修复：口语格式"都……了"（如"都十二点了"）限间隔 1-6 字且不跨标点；
    # 并排除"全都/全部+都"（范围副词"都"标记，如"禾苗全都死了"）
    if regex == "都.+了":
        regex = "(?<!全)(?<!全部)都[^，。]{1,6}了"
    return regex


def _grammar_name(content, det, category):
    """P1-4 语法点名称规范化：优先含格式标记(A/B/……)的内容 → 细目 → 类别+内容；绝不用正则。"""
    if content and ("A" in content or "……" in content):
        return content[:30]
    if det:
        return det
    if content:
        return content[:20]
    return category


# P4-1：官方语法表部分三级/高等语法点正则缺失，以下为补全正则（用于覆盖率自动初筛；
# 复杂句法/语义点仍标为“需人工复核”）。
GRAMMAR_REGEX_SUPPLEMENT = {
    "23": r"[0-9]+(个|只|条|张|本|支|块|件|双|岁|年|天|次|遍|口|位|名|种|些|点)",
    "24": r"(我|你|他|她|它|我们|你们|他们|她们|这|那|谁|什么|中国|人|学生|老师|爸爸|妈妈|朋友|孩子|时间|地方|东西)[^，。！？；]{1,15}(是|有|在|去|来|说|看|吃|喝|喜欢|爱|工作|学习|帮助|讨论)",
    "25": r"(是|有|在|去|来|说|看|吃|喝|喜欢|爱|觉得|知道|工作|学习|帮助|介绍|讨论)",
    "26": r"(是|有|在|去|来|说|看|吃|喝|喜欢|爱|觉得|知道|买|卖|给|送)[^，。！？；]{1,15}",
    "27": r"[\u4e00-\u9fff]{1,12}的[\u4e00-\u9fff]{1,12}",
    "28": r"(很|非常|都|也|在|正在|已经|马上|常常|经常|一起|跟|和|把|被)[^，。！？；]{1,15}",
    "29": r"(了|着|过|在|去|来|说|看|吃|喝|喜欢|爱|学习|工作|帮助|讨论)",
    "30": r"(很|非常|太|真|好|大|小|高|矮|漂亮|美丽|干净|热闹|冷|热|忙|累|快乐|高兴|幸福)",
    "31": r"(下雨|刮风|下雪|请|别|禁止|谢谢|你好|再见|对不起)",
    "32": r"。",
    "33": r"[？?]|吗|呢|谁|什么|哪|怎么|几|多少",
    "34": r"(请|别|不要|禁止|要|让|快|来)",
    "35": r"[！!]",
    "38": r"(比|没有)[^，。！？；]{1,20}",
    "40": r"了",
    "41": r"了",
    "43": r"[0-9]+(元|块|角|毛|分)",
    "44": r"[0-9]{1,4}(年|月|日|号)|(星期|周)[一二三四五六日天]|[0-9]{1,2}(点|时)[0-9]{0,2}分?",
    "48": r"(是不是|能不能|有没有|去不去|来不来|吃不吃|喝不喝|看不看|听不听|要不要|想不想|喜不喜欢|好不好|对不对|行不行|可不可以)",
    "85": r"(和|跟|与|或|或者|并|并且|也|都|的)",
    "86": r"[\u4e00-\u9fff]{1,10}的",
    "87": r"[\u4e00-\u9fff]{1,12}的[\u4e00-\u9fff]{1,12}",
    "88": r"(在|正在|去|来|看|说|吃|喝|学习|工作|帮助|讨论|介绍)[\u4e00-\u9fff]{1,12}",
    "89": r"(很|非常|太|真|挺|比较|那么|这么|多么)[\u4e00-\u9fff]{1,10}",
    "96": r"(今天|明天|昨天|现在|今年|明年|去年|这|那)[^，。！？；]{1,15}",
    "99": r"得[\u4e00-\u9fff]{1,10}",
    "100": r"[0-9]+(次|遍|下|回|趟|场)",
    "101": r"(大|小|长|短|高|矮|多|少|重|轻|早|晚)[0-9]+(岁|米|斤|厘米|倍|年|天|小时|分钟)",
    "102": r"(今天|明天|昨天|现在|今年|明年|去年|这|那)[^，。！？；]{1,15}",
    "103": r"有",
    "104": r"(处|边|上|下|里|外|前|后)[^，。！？；]{0,10}(有|着)",
    "105": r"(去|来|到|回|进|出)[^，。！？；]{1,15}(去|来|到|回|进|出|看|买|吃|喝|玩|学习|工作)",
    "109": r"(给|送|告诉|教|问|借|还|递|卖|买)[^，。！？；]{1,20}",
    "118": r"着",
    "119": r"过",
    "120": r"(第[0-9一二三四五六七八九十]+|[0-9]+(年|月|日|号|期|名|楼|层))",
    "121": r"[0-9]+多(岁|米|块|元|个|天|年|小时|次|斤)|[0-9]+(岁|米|块|元|个|天|年|小时|次|斤)多",
    "122": r"就",
    "131": r"[\u4e00-\u9fff]{2,8}(儿|家|们|头|子)",
    "136": r"(谁|什么|哪儿|哪里|怎么|怎样|多少|几)[^，。！？；]{0,10}(都|也)",
    "140": r"(个个|天天|年年|次次|回回|声声|张张|条条|件件|样样|句句|步步|人人|家家|户户|处处|顿顿)",
    "154": r"(因为|由于|为)[^，。！？；]{1,20}",
    "160": r"哈哈",
    "161": r"(在|从|对|把|向|跟|给|关于|对于|为了|除了)[^，。！？；]{1,15}",
    "162": r"[0-9]+[^，。！？；]{1,8}[0-9]+",
    "172": r"(我|你|他|她|它|我们|你们|他们|她们|这|那|谁|什么|中国|人|学生|老师|孩子|朋友|时间|地方)[^，。！？；]{1,15}(是|有|在|去|来|说|看|吃|喝|喜欢|爱|工作|学习|帮助|讨论)",
    "173": r"(是|有|在|去|来|说|看|吃|喝|喜欢|爱|觉得|知道|认为|希望|发现|表示)[^，。！？；]{1,15}",
    "174": r"[\u4e00-\u9fff]{1,12}的[\u4e00-\u9fff]{1,12}",
    "175": r"(到|住|走|完|好|懂|见|开|上|下|出)",
    "176": r"(出来|出去|过来|过去|回来|回去|进来|进去|起来|上来|上去|下来|下去)",
    "177": r"(得|不)(了|到|起|下|上|完|住|动|走|回|来|去)",
    "178": r"(得很|极了|死了)",
    "179": r"[0-9]+(次|遍|场|下|回|趟)",
    "180": r"[0-9]+(年|天|小时|分钟|个月|星期|周)",
    "181": r"[0-9]+(年|天|小时|星期|周|月)(以后|之后|以前|之前|来|以来)",
    "182": r"(这|那|他|她|它|我|你|我们|你们|他们|她们)[^，。！？；]{1,15}(是|有|在|去|来|说|看|吃|喝|喜欢|爱|工作|学习|帮助|讨论)",
    "183": r"把",
    "184": r"(被|叫|让|给)",
    "185": r"(去|来|到|回|进|出|用|拿|带|开|骑|坐)[^，。！？；]{1,15}(去|来|到|回|进|出|看|买|吃|喝|玩|学习|工作|说|写|画)",
    "186": r"(叫|派|请|让|使|令)[^，。！？；]{1,15}",
    "187": r"比",
    "188": r"(走一走|看一看|想一想|笑了笑|尝了尝|试试|研究研究|说说|聊聊|谈谈|听听)",
    "190": r"一会儿[^，。！？；]{1,20}一会儿",
    "192": r"又[^，。！？；]{1,12}又",
    "193": r"首先[^，。！？；]{1,30}然后",
    "195": r"(不仅|不光)[^，。！？；]{1,30}(还|而且)",
    "196": r"不是[^，。！？；]{1,30}就是",
    "197": r"是[^，。！？；]{1,12}就是",
    "198": r"要是[^，。！？；]{1,30}就",
    "199": r"只有[^，。！？；]{1,20}才",
    "200": r"(由于[^，。！？；]{1,20}(所以|因此))|([^，。！？；]{1,20}所以)",
    "201": r"为了[^，。！？；]{1,20}",
    "202": r"了[^，。！？；]{1,12}就",
    "203": r"(大概|大约|左右|前后)|([0-9]+[0-9]+)",
    "206": r"是[^，。！？；]{1,15}的",
    "207": r"[？?]",
    "208": r"都[^，。！？；]{1,12}了",
    "232": r"啊",
    "239": r"一[^，。！？；]{1,8}比一[^，。！？；]{1,8}",
    "244": r"(我|你|他|她|它|我们|你们|他们|她们|这|那|谁|什么|中国|人|学生|老师|孩子|朋友)[^，。！？；]{1,15}(是|有|在|去|来|说|看|吃|喝|喜欢|爱|工作|学习|帮助|讨论)",
    "245": r"(被|让|叫|给)[^，。！？；]{1,15}",
    "246": r"[\u4e00-\u9fff]{1,12}的[\u4e00-\u9fff]{1,12}的[\u4e00-\u9fff]{1,12}",
    "247": r"(上|出|起|下)(?=[，。！？；]|$)",
    "248": r"把",
    "249": r"被",
    "250": r"(出来|出去|进来|进去|回来|回去|过来|过去|起来|下来|下去)",
    "251": r"(表扬|批评|叫|称|说|收|选)[^，。！？；]{1,20}",
    "252": r"是[^，。！？；]{1,15}的",
    "254": r"既[^，。！？；]{1,20}(又|也)",
    "256": r"于是",
    "257": r"甚至",
    "258": r"或者[^，。！？；]{1,20}或者",
    "259": r"然而",
    "262": r"万一[^，。！？；]{1,20}就",
    "264": r"无论[^，。！？；]{1,20}(都|也)",
    "266": r"可见",
    "267": r"哪怕[^，。！？；]{1,20}(也|还)",
    "268": r"，好[^，。！？；]{1,10}",
    "269": r"[\u4e00-\u9fff]{1,8}(就|也|才|都)[\u4e00-\u9fff]{1,8}",
    "270": r"不[^，。！？；]{1,12}也",
    "271": r"[0-9]+来(岁|米|块|元|个|天|年|小时|次|斤)",
    "273": r"(谁|什么|哪儿|哪里|怎么|为什么|几|多少)[^，。！？；]{1,20}",
    "274": r"(不|没)[^，。！？；]{1,12}(不|没)",
    "312": r"从[^，。！？；]{1,20}来看",
    "318": r"(是|有|在|到|成|为|作|成为)[^，。！？；]{1,15}",
    "319": r"(很|非常|都|也|正在|已经|马上|常常|经常|一起|跟|和|在|把|被)[^，。！？；]{1,15}(很|非常|都|也|正在|已经|马上|常常|经常|一起|跟|和|在|把|被)",
    "320": r"(起来|下去|下来|上)",
    "321": r"(得|不得)(?=[，。！？；]|$)",
    "322": r"(不得了|慌|厉害|坏|透)了?",
    "323": r"得[^，。！？；]{1,20}",
    "324": r"(有着|有)",
    "325": r"把",
    "326": r"(卖|写|画|做|洗|修|改|盖|建|出版|翻译)",
    "327": r"(去|来|到|回|进|出|用|拿|带|开|骑|坐)[^，。！？；]{1,20}",
    "328": r"(叫|令|使|让)[^，。！？；]{1,20}",
    "329": r"(跟|和|与|同)[^，。！？；]{1,20}相比|比[^，。！？；]{1,20}",
    "330": r"或是[^，。！？；]{1,20}或是",
    "331": r"尽管[^，。！？；]{1,20}(但是|可是)",
    "333": r"要是[^，。！？；]{1,20}否则",
    "335": r"除非[^，。！？；]{1,20}(否则|不然)",
    "336": r"因而",
    "337": r"即使[^，。！？；]{1,20}也",
    "339": r"以便",
    "341": r"再[^，。！？；]{1,12}也",
    "342": r"[；;]",
    "343": r"(再也不|再没)",
    "344": r"可(好|美|真|太|别|不|没|想|喜欢|有意思|行)",
    "356": r"(他|她|它|他们|她们|它们|这|那|这些|那些|其|该)[^，。！？；]{1,15}",
    "357": r"(，)[^，。！？；]{1,12}(，)",
    "358": r"(超|多|反|无|亚|准)[\u4e00-\u9fff]{2,8}",
    "359": r"[\u4e00-\u9fff]{2,8}(化|式|型|性)",
    "381": r"[0-9]+(大|小|长|短|高|矮|厚|薄|重|轻|宽|窄)(个|只|条|张|块|件|片|间|座|层|倍)",
    "392": r"[^，。！？；]{1,6}一(个|只|条|张|块|件|片|间|座|层|倍)[^，。！？；]{1,6}一(个|只|条|张|块|件|片|间|座|层|倍)",
    "395": r"(到|去|在|成|为|作|成为|做出|得出)[^，。！？；]{1,15}",
    "396": r"(下来|下去|起来|过来)",
    "397": r"把",
    "398": r"被[^，。！？；]{1,20}给",
    "400": r"一时[^，。！？；]{1,20}一时",
    "401": r"便",
    "404": r"连[^，。！？；]{1,20}(都|也)[^，。！？；]{1,20}更",
    "405": r"要么[^，。！？；]{1,20}要么",
    "406": r"虽[^，。！？；]{1,20}(但|可|却|也)",
    "407": r"(要不然|不然)",
    "408": r"凡是[^，。！？；]{1,20}都",
    "409": r"(就算|就是)[^，。！？；]{1,20}也",
    "410": r"不[^，。！？；]{1,12}不",
    "411": r"[；;]",
    "412": r"非[^，。！？；]{1,15}不可",
    "434": r"[0-9]+(种|类|项|个|门|套|系列|方面)",
    "456": r"(用|拿|靠|通过|作为|为了|以)[^，。！？；]{1,15}",
    "458": r"个(不停|没完|够|遍|痛快|正着|满怀|措手不及|团团转|水泄不通)",
    "459": r"把",
    # V41：497/515/516 原补全用纯标点（；/。）作“触发正则”，几乎每句必命中，
    # 导致候选触发统计饱和（如“候选触发 72/81”信息量稀释）——已移除；
    # 这三条仍会显示在规则状态表（无正则/无检测器 → 暂不支持），供人工识别。
    "514": r"(因为|所以|但是|可是|于是|然后|并且|而且|如果|虽然|尽管|只要|只有|无论|即使|哪怕|除非|要么|或者|不但|不仅|与其|宁可|因此|因而|从而|接着|同时|然而|不过|否则|不然|何况|况且|再说|总之|例如|比如|也就是说)",
    "517": r"(因为|所以|因此|但是|可是|如果|只要|只有|为了|以免|否则|虽然|尽管|即使|哪怕)",
}

# 上述补全中需要按“句式/整句正则”匹配的编号（原表误标为“词类”或按词匹配过弱）
GRAMMAR_SENTENCE_SUPPLEMENT_NOS = {
    "100", "101", "102", "103", "104", "105",
    "109", "118", "119", "120", "121", "122",
    "131", "136", "140", "154", "160",
    "161", "162", "172", "173", "174", "175",
    "176", "177", "178", "179", "180", "181",
    "182", "183", "184", "185", "186", "187",
    "188", "190", "192", "193", "195", "196",
    "197", "198", "199", "200", "201", "202",
    "203", "206", "207", "208", "23", "24",
    "239", "244", "245", "246", "247", "248",
    "25", "249", "250", "251", "252", "254",
    "256", "257", "258", "26", "259", "262",
    "264", "266", "267", "268", "27", "269",
    "270", "271", "273", "274", "28", "29",
    "30", "31", "312", "318", "32", "319",
    "320", "321", "322", "323", "324", "325",
    "326", "327", "328", "33", "329", "330",
    "331", "333", "335", "336", "337", "34",
    "339", "341", "342", "343", "344", "35",
    "356", "357", "358", "359", "38", "381",
    "392", "395", "396", "397", "398", "40",
    "400", "401", "404", "405", "406", "407",
    "408", "41", "409", "410", "411", "412",
    "43", "434", "44", "456", "458", "459",
    "48", "514", "517",
    "85", "86", "87", "88", "89", "96",
    "99",
}


# 非索引分支（_GRAM_INDEX_USE=False）用于兼容旧路径的禁用编号集合。
# 当前版本默认走索引路径，此类编号已通过“补全正则只做候选、不计可信覆盖”处理。
GRAMMAR_COVERAGE_DISABLED_NOS = set()

# 候选规则升级为可信规则的验收阈值（修订版口径）：
# 至少抽查 GRAMMAR_TRUSTED_MIN_SAMPLE 个真实命中点，错误不超过 GRAMMAR_TRUSTED_MAX_ERRORS 个，
# 即最低精确率 GRAMMAR_TRUSTED_MIN_PRECISION（10个命中点最多错1个）。
# 【现状（V22）】：以下阈值为后续人工抽样预留，暂未参与 trusted 判定——
# 当前 trusted 仍由 load_gf0025 的 "source == official and no not in GRAMMAR_UNRELIABLE_NOS"
# 决定（约 400 条规则仅凭官方表可信）。若要启用证据门控，需引入
# GRAMMAR_RULE_EVIDENCE（编号→(样本数,错误数)）并在 load_gf0025 中按阈值裁决。
GRAMMAR_TRUSTED_MIN_PRECISION = 0.90
GRAMMAR_TRUSTED_MIN_SAMPLE = 10
GRAMMAR_TRUSTED_MAX_ERRORS = 1

# V55 金标门控：必须通过金标证据验收才能维持"可信"的高风险规则（用户评审结论：
# "不能把所有'可信命中'直接当作人工标注结果……至少编号22、33、82、108必须复核"）。
# 这些规则当前均无金标证据（data/语法规则证据.json 仅覆盖 183/184/186/187），
# 按"宁可让人工复核，也不让高风险规则直接进入公式"的策略 → 无证据即降为候选。
# 待补足 >= GRAMMAR_TRUSTED_MIN_SAMPLE 条且 errors <= GRAMMAR_TRUSTED_MAX_ERRORS 后自动恢复可信。
GRAMMAR_REVIEW_REQUIRED_NOS = {"22", "33", "82", "108"}

# 可信规则样本验收证据（第6项）：编号 -> {"samples": n, "errors": m, "tp/fp/fn": ...}。
# 数据源：data/语法规则证据.json（由 build_grammar_evidence.py 依据
# data/语法金标样本.csv 生成，需 LTP 实跑）。无证据编号回退“配置身份”（未验收）；
# 有证据但 samples<10 或 errors>1 → _evidence_ok=False → 自动降级为候选。
GRAMMAR_RULE_EVIDENCE = {}
_RULE_EVIDENCE_LOADED = False


def _get_rule_evidence():
    """lazy 加载语法规则证据（编号 -> 样本统计）；文件缺失返回 {}。"""
    global _RULE_EVIDENCE_LOADED
    if not _RULE_EVIDENCE_LOADED:
        _RULE_EVIDENCE_LOADED = True
        p = os.path.join(DATA_DIR, "语法规则证据.json")
        if os.path.exists(p):
            try:
                with open(p, encoding="utf-8") as f:
                    GRAMMAR_RULE_EVIDENCE.update(json.load(f))
            except Exception:
                pass
    return GRAMMAR_RULE_EVIDENCE


def _evidence_ok(no):
    """第6项/V55：样本验收门控。高风险规则（GRAMMAR_REVIEW_REQUIRED_NOS）无证据 → False
    （降级候选，必须人工复核）；其余规则无证据回退 True=配置身份；有证据但
    samples<10 或 errors>1 → False 降级候选。"""
    ev = _get_rule_evidence().get(no)
    if not ev:
        return no not in GRAMMAR_REVIEW_REQUIRED_NOS
    samples = ev.get("samples", 0) if isinstance(ev, dict) else ev[0]
    errors = ev.get("errors", 0) if isinstance(ev, dict) else ev[1]
    return samples >= GRAMMAR_TRUSTED_MIN_SAMPLE and errors <= GRAMMAR_TRUSTED_MAX_ERRORS

# 官方表中部分正则过宽、高级义项无法可靠识别（如“还/就”语气副词、
# “在……上/下/中”“以……为……”“什么X的Y的”等）；仅计入候选，不计入可信覆盖。
# 【编号对齐说明】编号以 data/GF0025_语法_with_regex.csv（删脚注重编号版，1..572）为准；
# 本集合曾于提交 b6d56d4 被整体误减 1（222→221、243→242、283→282……），
# 而数据表未同步变化，导致黑名单指向了错误规则（如“在……方面”），
# 本应降为候选的“在……上/下/中”“什么X的Y的”等重新被当作可信规则。
# 以下列表已按数据表逐条核对：
# 222 语气副词（的确/反而/还/竟然/究竟）　243 在……上/下/中　283 还X呢
# 294 语气副词（毕竟/不免/倒是/干脆/居然/可/明明/总算等）　303 “来X去”
# 309 不敢当　311 用不着　323 状态补语（无正则）　404 递进复句（无正则）
# 421 放着X不Y　454 以……为……　455 因……而……　475 ……，而……（则）……
# 481 任……，也……　492 ……，以……　505 不知X的　509 哪有X这么Y的
# 510 什么X的Y的
GRAMMAR_UNRELIABLE_NOS = {
    "222", "243", "283", "294", "303", "309", "311", "323",
    "404", "421", "454", "455", "475", "481", "492", "505",
    "509", "510",
    # V50（A2）：高频不挑句/宽泛误报的"补语/紧缩/前缀口语格式"规则 → 降候选。
    # 依据 precision_matrix.py（225 课）命中：98 趋向补语(上/来/起/去/过)3157、
    # 130 前缀(第-/老-/小-)命中"老师/小时/老人/老电影"1155 去重、97 结果补语
    # (好/会/完)1497、117"一……就……"285、129"要……了"91、210"X什么啊"281——
    # 这些规则缺句内结构校验，把"老师/小时/第一"等词面误判为语法点；降为候选
    # 供人工复核（不再进入可信覆盖与难度分值）。后续可仿 56/52 用 LTP 词性/结构窗口重构。
    "97", "98", "117", "129", "130", "210",
    # V53（高级语法误判，用户评审第四点）：够……的(314)、倒/反倒(476)、
    # 因……故(482)、说到/想到哪儿去了(511) —— 句式过宽、命中多为字面误报，
    # 降候选供人工复核（不进可信覆盖/分值）。
    "314", "476", "482", "511",
}


def load_gf0025():
    char_file = os.path.join(DATA_DIR, "GF0025_汉字.csv")
    # 优先使用“表面形式别名表”（由规范条目表展开生成）；缺失时回退旧运行时表
    alias_file = os.path.join(DATA_DIR, "GF0025_词汇_别名.csv")
    word_file = alias_file if os.path.exists(alias_file) else os.path.join(DATA_DIR, "GF0025_词汇.csv")
    gram_file = os.path.join(DATA_DIR, "GF0025_语法_with_regex.csv")

    char_levels = {}
    with open(char_file, encoding="utf-8") as f:
        next(f)
        for row in csv.reader(f):
            if len(row) >= 3:
                lv = _parse_level(row[1].strip())
                if lv:
                    char_levels[row[2].strip()] = lv

    word_levels = {}
    word_pos_levels = defaultdict(list)  # (词形, 词性大类) -> [等级]
    alias_mode = os.path.basename(word_file) == "GF0025_词汇_别名.csv"
    with open(word_file, encoding="utf-8") as f:
        next(f)
        for row in csv.reader(f):
            if alias_mode:
                if len(row) < 6:
                    continue
                lv = _parse_level(row[1].strip())
                w = row[3].strip()
                pos_field = row[5].strip()
            else:
                if len(row) < 3:
                    continue
                lv = _parse_level(row[1].strip())
                w = row[2].strip()
                pos_field = row[4].strip() if len(row) >= 5 else ""
            if lv and w:
                # 一词多级时取较低级别（保守口径）；同时保留 (词形, 词性) 维度
                if w not in word_levels or lv < word_levels[w]:
                    word_levels[w] = lv
                cats = [c.strip() for c in pos_field.split("、") if c.strip()]
                if not cats:
                    cats = [""]
                for cat in cats:
                    levels = word_pos_levels.setdefault((w, cat), [])
                    if lv not in levels:
                        levels.append(lv)

    grammar_rules = []
    with open(gram_file, encoding="utf-8") as f:
        header = next(f).strip().split(",")
        idx = {name: i for i, name in enumerate(header)}
        for row in csv.reader(f):
            if len(row) <= idx["正则表达式"]:
                continue
            no = row[0].strip() if row and row[0].strip() else ""
            regex = row[idx["正则表达式"]].strip()
            content = row[idx["语法内容"]].strip() if "语法内容" in idx and len(row) > idx["语法内容"] else ""
            det = row[idx["细目"]].strip() if "细目" in idx and len(row) > idx["细目"] else ""
            category = row[idx["类别"]].strip()
            project = row[idx["语法项目"]].strip()
            source = "official"
            # 精确语法规则策略：
            # - 官方表自带正则 → 可信自动检测（trusted=True）
            # - 官方表缺失正则、仅有补全正则 → 仅作候选提示（trusted=False），
            #   不计入“可自动检测/可信覆盖”，避免为了覆盖率而扩张规则。
            if no in GRAMMAR_REGEX_SUPPLEMENT:
                regex = GRAMMAR_REGEX_SUPPLEMENT[no]
                source = "supplement"
                if no in GRAMMAR_SENTENCE_SUPPLEMENT_NOS:
                    project = "补全句式"
            if not regex:
                continue
            # P1-3 后缀类规则（-化/-子/-儿/-们 等）：词尾正则无法区分"文化"与"现代化"，
            # 语言学上不可靠（审查指出"不能只因包含化便判为后缀"）→ 禁用匹配，报告注明人工复核
            if ("后缀" in det or det.startswith("-")) and source != "supplement":
                continue
            regex = _fix_format_regex(regex, content, det)  # P1-2 四字格单字紧邻
            # C-修复：记录标准语法点编号（官方表第一列 No.），用于"覆盖统计"
            no = row[0].strip() if row and row[0].strip() else ""
            # P1-4：口语格式"都……了"正则无法区分"都十二点了"(程度/时间到达) 与
            # "禾苗都死了"(范围副词+状态变化) → 设为"需人工复核"，不计自动覆盖
            if "都……了" in content or regex == "(?<!全)(?<!全部)都[^，。]{1,6}了":
                no = ""
            # P1-5：四字格"不A不B"正则无法区分"不知不觉"类成语与"不得不/不仅不/不给不"
            # 等短语 → 设为"需人工复核"，不计自动覆盖（与论文结论无关，纯规则能力限制）
            if "不A不B" in content:
                no = ""
            # B-修复："是……的"强调句需词性窗口（区间内须含动词，排除判断句"庄子是…哲学家"）
            pos_window = content.startswith("“是……的”句") or "是……的" in content
            grammar_rules.append({
                "no": no,
                "level": _parse_level(row[idx["级别"]].strip()),
                "project": project,
                "category": category,
                "det": det,
                "name": _grammar_name(content, det, category),  # P1-4 名称规范化
                "regex": regex,
                "pos_window": pos_window,
                "trusted": source == "official" and no not in GRAMMAR_UNRELIABLE_NOS
                           and _evidence_ok(no),
            })
    return char_levels, word_levels, grammar_rules, dict(word_pos_levels)


def load_char_freq():
    """字频表 -> {char: 每百万频率}

    数据源：data/character_frequency.csv（第 4 列"ferquency(pre million)"）。
    母语公式依赖本表；若文件缺失则返回 {}，analyze_native 中 avg_char_freq=0、
    公式的 -45.65*avg_char_freq/1e6 项静默为 0，母语分值系统性偏高
    （实测同一文本 40.5537 vs 40.3332）。缺失仅提示于 docstring，不主动告警。
    """
    freq = {}
    path = os.path.join(DATA_DIR, "character_frequency.csv")
    if not os.path.exists(path):
        return freq
    with open(path, encoding="utf-8") as f:
        next(f)
        for row in csv.reader(f):
            if len(row) >= 4 and row[1].strip():
                try:
                    freq[row[1].strip()] = float(row[3].strip())
                except ValueError:
                    pass
    return freq


# ---------------------------------------------------------------------------
# 分句 / 统计（纯汉字口径 v3：只统计汉字，不含标点/数字/字母）
# ---------------------------------------------------------------------------
def split_sentences(text):
    """切分句子，并把句末引号（“ ” ‘ ’ 等）归并到前一句，避免出现单独的引号“句子”。"""
    parts = [s for s in re.split(r"(?<=[。！？；!?;\n])", text) if s.strip()]
    out = []
    for part in parts:
        s = part.strip()
        if out:
            # 纯引号/空白段 → 并入前句（如“...快乐啊！”后面的引号）
            if re.fullmatch(r"[\s“”\"'‘’]+", s):
                out[-1] = out[-1] + s
                continue
            # 段首的闭合引号 → 移回前句末（如“...呢？”庄子...）
            m = re.match(r"^[”\"’]+", s)
            if m:
                out[-1] = out[-1] + m.group(0)
                s = s[m.end():].strip()
                if not s:
                    continue
        out.append(s)
    return [s for s in out if s.strip()]


# 超长单句切分时允许的“回退搜索边界”距离（在 [max_chars*0.5, max_chars] 内找标点/空白）
_LONG_SENT_MIN_RATIO = 0.5
# GF0025 词表最长词为 4 字（实测），且所有标准词 ≤4 字；边界前瞻用 4 字即可覆盖
# 全量 4 字标准词（含“爱不释手/持之以恒/阿拉伯语”等 281 个边界词）。
_STD_MAX_WORD_LEN = 4
_STD_PREFIX_SET = None  # 惰性构建：标准词真前缀集合（切点安全校验用）
_STD_WORD_SET = None    # 惰性构建：标准词完整集合


def _get_std_prefix_set():
    """返回 GF0025 词表全部标准词的【真前缀+完整词】集合（惰性构建）。
    用于判断切点是否落在标准词内部（避免切断 4 字词）。
    """
    global _STD_PREFIX_SET, _STD_WORD_SET
    if _STD_PREFIX_SET is None:
        _cl, word_levels, _gr = _get_gf0025()
        prefixes = set()
        full = set()
        for w in word_levels:
            if len(w) > _STD_MAX_WORD_LEN:
                continue  # 理论上无 >4 字词；保守跳过
            full.add(w)
            for k in range(1, len(w)):
                prefixes.add(w[:k])
        _STD_PREFIX_SET = prefixes
        _STD_WORD_SET = full
    return _STD_PREFIX_SET


def _get_std_word_set():
    """返回 GF0025 词表完整标准词集合（惰性构建）。"""
    _get_std_prefix_set()
    return _STD_WORD_SET or set()


def _long_sentence_chunks(sent, max_chars):
    """把超过 max_chars 的单句切分成若干 ≤max_chars 的块。

    - 优先在标点/空白处切分（不截断词语）；
    - 找不到标点边界时，用 LTP 词边界切分，并用 GF0025 词表前缀集合
      校验切点不会切断任何标准词（见 _word_boundary_cut/_cut_is_safe）；
    返回块列表；块之间无重叠，拼接后完整覆盖原句。
    """
    chunks = []
    rest = sent
    while len(rest) > max_chars:
        low = int(max_chars * _LONG_SENT_MIN_RATIO)
        cut = -1
        # 1) 标点/空白边界：在 [low, max_chars] 区间内找最后一次出现
        for i in range(max_chars, low - 1, -1):
            if i > len(rest):
                continue
            if rest[i - 1] in "，。！？；、！？\n\r\t " or rest[i - 1] in ".,;!? \t":
                cut = i
                break
        # 2) 无标点边界 → 用 LTP 词边界 + 词表前缀校验切分
        if cut <= 0 or cut >= len(rest):
            cut = _word_boundary_cut(rest, max_chars, low)
        if cut <= 0 or cut >= len(rest):
            # 兜底：从 max_chars 向前找“不切断标准词”的安全切点
            cut = _safe_fallback_cut(rest, max_chars, low)
        if cut <= 0 or cut >= len(rest):
            cut = max_chars  # 最终兜底：硬切（全字母/数字无词边界时）
        chunks.append(rest[:cut])
        rest = rest[cut:]
    if rest.strip():
        chunks.append(rest)
    return chunks or [sent[:max_chars]]


def _safe_fallback_cut(sent, max_chars, low):
    """无 LTP 词结束位置可用时，从 max_chars 向前找“不切断标准词”的切点。"""
    std_prefixes = _get_std_prefix_set()
    for i in range(max_chars, low - 1, -1):
        if i > len(sent):
            continue
        if _cut_is_safe(sent, i, std_prefixes):
            return i
    return -1


def _word_boundary_cut(sent, max_chars, low):
    """在 [low, max_chars] 内找安全切点：优先 LTP 分词词结束位置，
    且切点不能落在任何 GF0025 标准词内部（用词表真前缀集合校验）；
    找不到返回 -1。全项目分词统一 LTP（不再用 jieba）。
    """
    try:
        std_prefixes = _get_std_prefix_set()
        end = 0
        best = -1
        # 前瞻：多切 4 字（最长标准词），使跨边界的多字标准词能完整被 LTP 看到
        lookahead = sent[:max_chars + _STD_MAX_WORD_LEN]
        res = _ltp_analyze_chunks([lookahead])
        if not res:
            return -1
        cws = [w.strip() for w in res[0].get("cws", [])]
        for w in cws:
            end += len(w)
            if end <= max_chars and end >= low:
                # 切点安全校验：切点右侧第一个字符 + 左侧 ≤3 字符
                # 不能构成“某标准词的未完前缀”，否则会切断标准词。
                if _cut_is_safe(sent, end, std_prefixes):
                    best = end
        return best
    except Exception:
        return -1


def _cut_is_safe(sent, cut, std_prefixes, std_words=None):
    """切点 cut 是否安全：cut 位置不能位于任何标准词内部。

    std_prefixes：标准词真前缀集合（用于识别“某词从这里开始”）。
    std_words：标准词完整集合（用于判断 cand 是否仍是完整词）。
    若切点左侧 k 字符是某标准词的真前缀，且切点后字符使该词继续成立
    （left+nxt 是标准词前缀或完整词）→ 切点切断词，不安全。
    """
    if cut >= len(sent):
        return True
    nxt = sent[cut]
    std_words = std_words if std_words is not None else _get_std_word_set()
    for k in range(1, min(_STD_MAX_WORD_LEN, cut) + 1):
        left = sent[cut - k:cut]
        if left in std_prefixes:
            cand = left + nxt
            if cand in std_prefixes or cand in std_words:
                return False  # 词仍在延续（前缀/完整词），切点在词内部
    return True


def _chunk_text(text, max_chars=400):
    """LTP 输入块：正常句子【逐句】作为独立块（不合并句子），
    保持 LTP 分词上下文与逐句分析一致；仅当单个句子超过 max_chars 时
    才按边界感知切分/重叠窗口拆分（见 _long_sentence_chunks）。
    返回块列表。
    """
    chunks = []
    for sent in split_sentences(text):
        if not sent.strip():
            continue
        if len(sent) <= max_chars:
            chunks.append(sent)
        else:
            chunks.extend(_long_sentence_chunks(sent, max_chars))
    return chunks or [text]


def _sentence_list(text):
    """返回完整句子列表；每个句子作为一个 LTP 输入单位。"""
    return [s for s in split_sentences(text) if s.strip()] or [text]


def _is_hanzi(c):
    return "\u4e00" <= c <= "\u9fff"


def _hanzi_of(text):
    return [c for c in text if _is_hanzi(c)]


def local_stats(text):
    """统计口径：总字数=非空字符数；字种数=纯汉字；句长按非空字符计算。"""
    chars = [c for c in text if not c.isspace()]
    sents = split_sentences(text)
    sent_lens = [len([c for c in s if not c.isspace()]) for s in sents]
    return {
        "total_chars": len(chars),
        "unique_chars": len(set(c for c in chars if "一" <= c <= "鿿")),
        "sentences": len(sents),
        "avg_sent_len": round(len(chars) / len(sents), 1) if sents else 0,
        "max_sent_len": max(sent_lens, default=0),
    }


# ---------------------------------------------------------------------------
# 对外汉语模式：程勇 GF0025 公式 + 分级明细
# ---------------------------------------------------------------------------
_POS_MAP = {
    "副词": "d", "介词": "p", "助词": "u", "量词": "q", "代词": "r",
    "动词": "v", "连词": "c", "叹词": "e", "数词": "m", "形容词": "a", "名词": "n",
    "拟声词": "o",  # P0-2：LTP 词性 o（onomatopoeia）→ 拟声词
}

# V55：助词(u)规则消歧优先级。同一个 u 词可能匹配多条助词规则（20/21/22/79/80/81/82/380 等），
# V54 及之前会【全部记入】grammar_covered → 同一"的"算作 20（一级）+82（二级）2 种，
# 同一"了"算作 21（一级）+22（一级）2 种，语法种类被双计/双级别误报，难度公式的 r_grammar 分母被高估。
# 修复后：pos_index[u] 路径上对每个 u 词只归入优先级最高的 1 个 no，避免重复计分与跨级别污染。
# 优先级表键为 u 词的字面（"了/的/地/得/着/过/吧/吗/呢/啊/所/之/罢了/啦/嘛/而已/矣"），
# 值为该字面在 u-pos 下的【首选 no】。未列入的字面仍按原 regex 匹配结果（罕见词不受影响）。
_UHLP_DISAMBIG = {
    "了": "21",   # 了1 动态助词（绝大多数"了"是完成/动态体；"了2"语气助词仅在句末表变化/新情况）
    "的": "20",   # 的1 结构助词（绝大多数"的"是定语标记；"的2"语气助词仅在句末表肯定/已然）
    "地": "20",   # 地 结构助词
    "得": "79",   # 得 结构助词
    "着": "81",   # 着 动态/持续
    "过": "80",   # 过 动态/经历
    "吧": "22",   # 吧 语气助词（一级优先；二级 82 同义但等级低，保留上级）
    "吗": "22",   # 吗 语气助词
    "呢": "22",   # 呢 语气助词
    "啊": "82",   # 啊 语气助词
    "所": "379",  # 所 结构助词
    "之": "433",  # 之 结构助词
    "罢了": "380",
    "啦": "380",
    "嘛": "380",
    "而已": "571",
    "矣": "572",
}

# search 级词类规则（V48 修复）：56 恢复句级 search + 词性窗口 ("a","z")——
# LTP 将"高高的/高高兴兴"整词标 z（状态词），fullmatch 无法命中（词性组只认 a）；
# 表正则收紧为 `^(\S)\1的$|^(\S)\2(\S)\3$`（仅"AA的/AABB"，独立组号），
# 配合窗口 a/z 排除"常常/刚刚/往往/一一"（标 d/m），排除"谢谢/帮帮"（标 v）。
# 52（动词重叠）留 search + 窗口 v + 词表固定词排除（挡"谢谢常常"）；"走着走着"
# 属动词重叠持续体，接受。135（离合词 打不开/离不开）留 search（形态独特）。
_SEARCH_CLASS_NOS = {"52", "56", "135"}
_SEARCH_CLASS_POS = {"52": ("v",), "56": ("a", "z")}

_CHAR_CACHE, _WORD_CACHE, _GRAM_CACHE = None, None, None
_WORD_POS_LEVELS_CACHE = None
_GRAM_INDEX = None  # 语法规则索引：词性分组 + 预编译


def _get_gf0025():
    global _CHAR_CACHE, _WORD_CACHE, _GRAM_CACHE, _GRAM_INDEX, _WORD_POS_LEVELS_CACHE
    if _CHAR_CACHE is None:
        _CHAR_CACHE, _WORD_CACHE, _GRAM_CACHE, _WORD_POS_LEVELS_CACHE = load_gf0025()
        _GRAM_INDEX = build_grammar_index(_GRAM_CACHE)
    return _CHAR_CACHE, _WORD_CACHE, _GRAM_CACHE


def _get_word_pos_levels():
    """返回 (词形, 词性大类) -> [等级] 的联合词表（首次访问时加载）"""
    _get_gf0025()
    return _WORD_POS_LEVELS_CACHE or {}


def build_grammar_index(rules):
    """语法规则索引（性能优化 v2）：
    词类规则按词性分组（每词只查本词性规则组），非词类规则预编译正则。
    返回 (pos_index, sentence_rules)
    sentence_rules 元素: (compiled, level, name, no, pos_window)
    V45：_SEARCH_CLASS_NOS（52/56/135 形态可变词类）从词类组移入句级 search
    （去锚定），避免 LTP 拆词导致的整词 fullmatch 漏检。
    """
    from collections import defaultdict
    pos_index = defaultdict(list)          # POS -> [(compiled, level, name)]
    sentence_rules = []                    # [(compiled, level, name, no, pos_window)]
    for r in rules:
        if not r["level"]:
            continue
        name = r["name"]  # P1-4：name 已在 load 时规范化，不再用正则兜底
        regex = r["regex"]
        if r.get("no") in _SEARCH_CLASS_NOS and r["project"] == "词类":
            # 句级 search：去掉 ^/$ 锚（否则 search 等价于整句等值）
            regex = re.sub(r"[$^]", "", regex)
            r = dict(r)
            r["regex"] = regex
            # 词性窗口（字符串词性代码）：52=v、56=a；135 无窗口（形态独特）
            if r.get("no") in _SEARCH_CLASS_POS:
                r["pos_window"] = _SEARCH_CLASS_POS[r["no"]]
        try:
            creg = re.compile(r["regex"])
        except re.error:
            continue
        if r["project"] == "词类" and r.get("no") not in _SEARCH_CLASS_NOS:
            pos = _POS_MAP.get(r["category"])
            if pos:
                pos_index[pos].append((creg, r["level"], name, r.get("no", ""), r.get("trusted", True)))  # P0-1 词类规则也带编号
        else:
            sentence_rules.append((creg, r["level"], name, r.get("no", ""), r.get("pos_window", False), r.get("trusted", True)))
    return dict(pos_index), sentence_rules


def _get_grammar_index():
    if _GRAM_INDEX is None:
        _get_gf0025()
    return _GRAM_INDEX


_GRAMMAR_POINT_INDEX = None


def _grammar_point_index():
    """标准语法点编号 -> {"level": 1-7, "name": 名称, "det": 细目}（全量 572 条）。
    用于 LTP 依存句法检测器命中的编号映射到等级与名称。"""
    global _GRAMMAR_POINT_INDEX
    if _GRAMMAR_POINT_INDEX is not None:
        return _GRAMMAR_POINT_INDEX
    gram_file = os.path.join(DATA_DIR, "GF0025_语法_with_regex.csv")
    idx = {}
    try:
        with open(gram_file, encoding="utf-8") as f:
            header = next(f).strip().split(",")
            col = {n: i for i, n in enumerate(header)}
            for row in csv.reader(f):
                if len(row) <= col["级别"]:
                    continue
                lv = _parse_level(row[col["级别"]].strip())
                if not lv:
                    continue
                no = row[0].strip() if row and row[0].strip() else ""
                if not no:
                    continue
                content = row[col["语法内容"]].strip() if "语法内容" in col and len(row) > col["语法内容"] else ""
                det = row[col["细目"]].strip() if "细目" in col and len(row) > col["细目"] else ""
                category = row[col["类别"]].strip() if "类别" in col and len(row) > col["类别"] else ""
                idx[no] = {
                    "level": lv,
                    "name": _grammar_name(content, det, category),
                    "det": det,
                }
    except Exception:
        pass
    _GRAMMAR_POINT_INDEX = idx
    return _GRAMMAR_POINT_INDEX


def _grammar_catalog():
    """C-修复：官方语法点全量目录（按级别，含无正则/被禁用规则），供"覆盖统计"使用。
    直接读官方表，保证分母 = 《标准》全量（如三级 81 条），而非程序可匹配子集。
    返回 {lv: [(no, name, det)]}，lv 为 1-7。"""
    gram_file = os.path.join(DATA_DIR, "GF0025_语法_with_regex.csv")
    cat = {}
    try:
        with open(gram_file, encoding="utf-8") as f:
            header = next(f).strip().split(",")
            idx = {n: i for i, n in enumerate(header)}
            for row in csv.reader(f):
                if len(row) <= idx["级别"]:
                    continue
                lv = _parse_level(row[idx["级别"]].strip())
                if not lv:
                    continue
                no = row[0].strip() if row and row[0].strip() else ""
                content = row[idx["语法内容"]].strip() if "语法内容" in idx and len(row) > idx["语法内容"] else ""
                det = row[idx["细目"]].strip() if "细目" in idx and len(row) > idx["细目"] else ""
                category = row[idx["类别"]].strip() if "类别" in idx and len(row) > idx["类别"] else ""
                cat.setdefault(lv, []).append(
                    (no, _grammar_name(content, det, category), det))
    except Exception:
        pass
    for lv in cat:
        cat[lv].sort(key=lambda x: x[0])
    return cat


# 语法匹配模式开关：True=词性索引+预编译（性能优化版） / False=全量遍历（普通版）
_GRAM_INDEX_USE = True


def gf0025_level(score):
    """GF0025 分级映射（程勇2023表8【论文原始边界】，回归 V43 前口径）。

    论文表8边界：[1-1.4)一级 [1.4-1.75)二级 [1.75-2)三级 [2-2.3)四级
    [2.3-2.55)五级 [2.55-2.7)六级 [2.7-7]七-九级。
    V44 曾用 no7z HSK 句子等级校准为 [1.25,1.6,1.85,2.15,2.4,2.85]，
    但 HSK 句级 ≠ 程勇文本难度级（不同变量），违反论文口径 ——
    V48 恢复论文原始边界。如需 HSK 校准模式，另做可选参数（勿改默认）。
    """
    for bound, name in ((1.4, "一级"), (1.75, "二级"), (2.0, "三级"),
                        (2.3, "四级"), (2.55, "五级"), (2.7, "六级")):
        if score < bound:
            return name
    return "七-九级"


def analyze_foreign(text):
    """对外汉语模式（生产入口统一 LTP）。

    本函数为兼容保留的代理：历史 Jieba 快速路径已移除，
    全项目计算只使用 LTP 分词/词性（见 analyze_foreign_ltp）。
    """
    return analyze_foreign_ltp(text)


# ---------------------------------------------------------------------------
# v4 完全本地模式：LTP 分词（与网站同一分词器）+ GF0025 公式
# ---------------------------------------------------------------------------
_LTP = None


def _ensure_model_merged(local_model):
    """若 LTP/small 只有 .bin 分片（partN），合并为完整 pytorch_model.bin。

    Gitee 不支持 git-lfs 且单文件限 100MB，模型按 90MB 分片入库；
    运行时先合并，保证 LTP 可正常加载。合并前校验：
    - 至少 2 个分片（1 个分片不足以组成大于 100MB 的模型）；
    - 每个分片非空；
    - 已存在的完整 bin 大小必须等于分片之和（截断/残留视为无效，重新合并）。
    写入采用【临时文件 + os.replace 原子替换】，避免残留部分写入。
    校验失败返回 False（不写坏文件）。
    """
    bin_path = os.path.join(local_model, "pytorch_model.bin")
    parts = sorted(
        p for p in os.listdir(local_model)
        if re.fullmatch(r"pytorch_model\.bin\.part\d+", p)
    )
    expect = sum(os.path.getsize(os.path.join(local_model, p)) for p in parts) if parts else 0
    if os.path.exists(bin_path):
        # 已有 bin：仅当大小等于分片之和且非零时视为有效；否则重新合并
        if expect > 0 and os.path.getsize(bin_path) == expect:
            return True
    if len(parts) < 2:
        return False  # 分片缺失/不足（至少需 2 片才能组成 >100MB 模型）
    for p in parts:
        if os.path.getsize(os.path.join(local_model, p)) <= 0:
            return False
    try:
        # 原子写入：先写临时文件，校验大小后再 os.replace 替换正式文件
        tmp_path = bin_path + ".tmp"
        with open(tmp_path, "wb") as out:
            for p in parts:
                with open(os.path.join(local_model, p), "rb") as pf:
                    shutil.copyfileobj(pf, out)
            out.flush()
            os.fsync(out.fileno())
        if os.path.getsize(tmp_path) != expect:
            os.remove(tmp_path)
            return False
        os.replace(tmp_path, bin_path)
        return True
    except OSError:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
        return False


def _get_ltp():
    """全局单例 LTP（延迟加载；打包后从 _MEIPASS/ltp_model 加载）"""
    global _LTP
    if _LTP is None:
        from ltp import LTP
        if getattr(sys, "frozen", False):
            _LTP = LTP(os.path.join(sys._MEIPASS, "ltp_model"))
        else:
            # 本地模型目录存在（LTP/small 含 config.json）→ 用绝对路径，
            # 避免从其他工作目录启动时找不到模型；
            # 分片模型（pytorch_model.bin.partN）先合并为 pytorch_model.bin；
            # 目录完全缺失（干净克隆/仅 HuggingFace 缓存）→ 回退模型名称，
            # 由 LTP/transformers 按缓存或默认下载逻辑加载。
            local_model = os.path.join(_HERE_DIR, "LTP", "small")
            if os.path.exists(os.path.join(local_model, "config.json")):
                if not os.path.exists(os.path.join(local_model, "pytorch_model.bin")):
                    _ensure_model_merged(local_model)
                _LTP = LTP(local_model)
            else:
                _LTP = LTP("LTP/small")
    return _LTP


def _ltp_analyze_chunks(chunks):
    """对文本块做 LTP 分析，返回 results 列表或 None。

    优先【进程内直接调用】LTP（打包 exe 与源码均可用：torch/ltp/transformers
    已随包内置或在本机环境）；进程内不可用（如本机未装 ltp）时，回退到
    ltp_bridge.py 子进程方式（需独立 Python 环境，设置 LTP_PYTHON 指向它）。
    """
    # 1) 进程内直接 LTP：单例 _get_ltp() 已处理 frozen(_MEIPASS/ltp_model)
    #    与源码(LTP/small)两种模型路径；按批调用避免整批内存过高。
    try:
        ltp = _get_ltp()
        results = []
        _BATCH = 16
        for i in range(0, len(chunks), _BATCH):
            batch = chunks[i:i + _BATCH]
            out = ltp.pipeline(batch, tasks=["cws", "pos", "dep"])
            for j in range(len(batch)):
                dep = out.dep[j]
                results.append({
                    "cws": out.cws[j],
                    "pos": out.pos[j],
                    "dep": {"head": [int(x) for x in dep["head"]], "label": dep["label"]},
                })
        return results
    except Exception as _exc:
        # 诊断：打印进程内 LTP 失败原因（打包环境定位用），随后走子进程桥接兜底
        import traceback as _tb
        print(f"[诊断] 进程内 LTP 调用失败:\n{_tb.format_exc()}", file=sys.stderr)
        pass  # 进程内不可用，走子进程桥接兜底

    # 2) 子进程桥接兜底：仅当用户显式设置 LTP_PYTHON（指向独立含 LTP 的
    #    Python 解释器）时才启用；打包后 sys.executable 是 exe 本身，
    #    不可作为 python 使用——否则会递归启动自身 exe，造成卡死/闪退。
    python = os.environ.get("LTP_PYTHON")
    if not python:
        return None
    import json as _json
    import subprocess as _sp
    here = os.path.dirname(os.path.abspath(__file__))
    bridge = os.path.join(here, "ltp_bridge.py")
    if not os.path.exists(bridge) or not os.path.exists(python):
        return None
    try:
        proc = _sp.run([python, bridge],
                       input=_json.dumps({"texts": chunks}, ensure_ascii=False),
                       capture_output=True, text=True, encoding="utf-8", timeout=600)
        if proc.returncode != 0:
            return None
        out = proc.stdout.strip()
        # 最后一行是 JSON，前几行可能是模型加载日志
        data = _json.loads(out.split("\n")[-1])
        return data.get("results", [])
    except Exception:
        return None


def _verify_ltp_coverage(chunks, ltp_results):
    """校验 LTP 输出是否完整覆盖输入块，防止长句/长块被 LTP 静默截断。

    - chunks 是生产入口实际送入 LTP 的文本块（已按 ≤400 字切块）；
    - ltp_results 是 LTP 返回结果（每块一个 dict，含 cws/pos/dep）。
    拼接各块 cws 得到的字符数与对应原块比较；LTP 分词会剔除少数特殊字符，
    因此阈值设为去空白后原块的 90%：低于阈值说明该块被截断/丢词，
    直接抛 RuntimeError（不静默），让调用方知晓输入未全部分析。
    返回 None；校验失败抛出异常。
    """
    if len(ltp_results) != len(chunks):
        raise RuntimeError(
            f"LTP 结果块数 {len(ltp_results)} 与输入块数 {len(chunks)} 不一致，"
            f"输入文本未被完整分析（可能被静默截断）。"
        )
    for i, (chunk, res) in enumerate(zip(chunks, ltp_results)):
        cws = res.get("cws") or []
        recovered = "".join(str(w) for w in cws)
        orig_len = len([c for c in chunk if not c.isspace()])
        recov_len = len([c for c in recovered if not c.isspace()])
        if orig_len and recov_len < orig_len * 0.9:
            raise RuntimeError(
                f"LTP 分析块 #{i + 1} 输出不完整：原块 {orig_len} 字符，"
                f"拼接词形仅 {recov_len} 字符（疑似被静默截断）。"
                f"请检查输入；若为超长句，应确保入口使用 _chunk_text 切块。"
            )
    return None


def _site_stats(text, words_per_sent, word_levels=None, poses_per_sent=None, word_pos_levels=None):
    """统计口径：总字数=非空字符数(与网站一致)；词数=LTP 分词（含标点词）。
    字种数=纯汉字（P2-1：过滤数字/字母/标点）；
    词种数：只统计【GF0025 官方词表内命中】的词（优先 (词形,词性) 联合定级，含词性兜底）。"""
    chars = [c for c in text if not c.isspace()]
    total_words = sum(len(ws) for ws in words_per_sent)
    # 纯汉字词数（不含标点/数字/字母 token），与字种/词种口径一致
    hanzi_words = 0
    for ws in words_per_sent:
        for w in ws:
            if re.fullmatch(r'[\u4e00-\u9fff]+', w):
                hanzi_words += 1
    uniq_words = set()
    _exact_seen_forms = set()  # 已精确命中/兜底记录过的词形，防同一词重复计词型
    for si, ws in enumerate(words_per_sent):
        for wi, w in enumerate(ws):
            if word_levels is not None:
                pos = ""
                if poses_per_sent and si < len(poses_per_sent) and wi < len(poses_per_sent[si]):
                    pos = poses_per_sent[si][wi]
                lv, _amb, src = _lookup_word_level_detail(word_levels, word_pos_levels or {}, w, pos)
                if not lv:
                    continue  # 不在官方词表内 → 不计入词种
                # 精确命中 (src=pos/empty) 用 (词形,词性) 键；词性兜底 (src=form)
                # 用【词形唯一键】，与 word_unique_counts 计分明细口径一致，
                # 避免同一词多词性被重复计数（如“很/副”+“很”算 2 个）。
                if src == "form":
                    if w not in _exact_seen_forms:
                        uniq_words.add(w)
                        _exact_seen_forms.add(w)  # 兜底词形也视为“已出现”，后续重复不计
                else:
                    uniq_words.add(_word_entry_key(w, pos))
                    _exact_seen_forms.add(w)
            else:
                uniq_words.add(w)
    sents = split_sentences(text)
    sent_lens = [len([c for c in s if not c.isspace()]) for s in sents]
    return {
        "total_chars": len(chars),
        "unique_chars": len(set(c for c in chars if "一" <= c <= "鿿")),  # P2-1 纯汉字
        "total_words": total_words,              # 兼容字段：分词 Token 数（含标点）
        "total_word_tokens": total_words,        # 明确语义：分词 Token 数（含标点）
        "total_hanzi_words": hanzi_words,        # 纯汉字词数（不含标点/数字/字母）
        "unique_words": len(uniq_words),
        "sentences": len(sents),
        "avg_sent_len": round(len(chars) / len(sents), 1) if sents else 0,
        "max_sent_len": max(sent_lens, default=0),
    }


# P1-1 凝固词表：四字格表面结构但为凝固词（"不A不B"等不应计入，如"不知不觉"）
_FROZEN_QUADS = {
    "不知不觉", "不得不说", "不吃不喝", "不是不", "不伦不类",
    "不三不四", "不痛不痒", "不上不下", "不闻不问", "不依不饶",
    "有声有色", "有说有笑", "有始有终", "有头有尾", "有来有往",
    "自以为是", "自作自受", "自高自大", "自暴自弃", "东张西望", "东拼西凑",
}


def _is_frozen_quad(matched_text):
    """P1-3：四字格命中是否为凝固词（应排除）。"""
    t = matched_text.strip()
    return any(w in t for w in _FROZEN_QUADS)


def _match_token_pos_in_span(match, words, poses, pos_prefix):
    """V48 词性窗口 + V54 形态排除 + V55 扩充排除 + V55 跨词匹配检测。

    匹配片段 [span) 覆盖的 token 中，至少一个词性以 pos_prefix 开头；且【整段匹配文本不在词表内】。

    V54 形态排除（52 号）：
      - "V着V着"（走着走着）——LTP 拆词后非标准 VV；
      - "一X一X"（一拐一拐）——数量/方式重叠。
    V55 扩充排除（52 号）：
      - "X啊X啊""X吧X吧"——语气助词重叠；
      - "X你X我"——代词对举；
      - "X来X回"——反义对举（在 chunk 内做上下文 ±2 字符检测，捕获"来来回回"被
        52 regex 拆为"来来""回回"两条的拆分情况）。

    V55 跨词匹配检测（56 号）：56 regex `(\\S)\\1的|(\\S)\\2(\\S)\\3` 在 search 模式下去锚后
    会从 LTP 完整 token 中间断开匹配（"珍珍的幸福"中匹配"珍珍的"，把"的"前的 token 拦腰切），
    V55 校验：匹配片段的起止位置必须落在 LTP token 边界上（与 token 起始或结束重合），
    否则视为跨词错误匹配。
    """
    start, end = match.span()
    prefixes = (pos_prefix,) if isinstance(pos_prefix, str) else tuple(pos_prefix)
    text = (match.group(0) or "")
    _wlv = _get_gf0025()[1]
    if _wlv and text.strip() in _wlv:
        return False  # 词表固定词：不接受为重叠语法
    if prefixes == ("v",):
        if re.search(r"[一-龥]着[一-龥]{0,2}着", text):
            return False  # V着V着（走着走着）
        if re.search(r"一[一-龥]{1,2}一[一-龥]{1,2}", text):
            return False  # 一X一X（一拐一拐/一摇一摇；不误伤"看一看"三字）
        # V55：四字口语/对举重叠排除
        if re.search(r"[一-龥]啊[一-龥]{0,1}啊", text):
            return False  # X啊X啊
        if re.search(r"[一-龥]吧[一-龥]{0,1}吧", text):
            return False  # X吧X吧
        if re.search(r"[你我他她它]你[一-龥]我", text):
            return False  # 你你我我/我你你他
    # V55 跨词匹配检测（仅对 56 号（prefix 含 a/z）做严格 token 边界校验）
    if "a" in prefixes or "z" in prefixes:
        # 计算 LTP token 边界位置
        token_boundaries = set()
        off = 0
        for w in words:
            w = w or ""
            if w:
                token_boundaries.add(off)
                token_boundaries.add(off + len(w))
            off += len(w)
        # 匹配片段的 start 和 end 必须都落在 token 边界上
        if start not in token_boundaries or end not in token_boundaries:
            return False
    off = 0
    chunk_text = ""  # V55：用于上下文对举检测
    for w, p in zip(words, poses):
        w = w or ""
        nxt = off + len(w)
        chunk_text += w
        if nxt > start and off < end:
            if any((p or "").startswith(pre) for pre in prefixes):
                # V55 上下文对举检测（仅在 prefix 是 v 时）
                if prefixes == ("v",):
                    ctx = chunk_text[max(0, start-2):min(len(chunk_text), end+2)]
                    if re.search(r"[一-龥]来[一-龥]回", ctx):
                        return False  # 来来回回等
                return True
        if nxt > end:
            break
        off = nxt
    return False


def _is_shi_de_emphasis(match, words, poses):
    """P1-1/V51/V55："是……的"强调句统一判定（所有运行模式共用）。

    用正则实际匹配区间（match.span）定位"是"与"的"，检查其间是否含【动词】词性
    （判断句"庄子是…哲学家"区间内全为名词 → 判否；强调句"他是昨天来的"含动词 → 判是）。
    V51 修正误报（原 530 次偏高）：
    - 改为【就近匹配】——取最后一个"是"与其后**第一个**以"的"结尾的词，
      而非"全句首是 → 全句末的"（后者会横跨大段、把相邻句语义混进来）；
    - 窗口词数 ≤6（过长视为非强调结构，判否）；
    - 排除"是"前紧邻"不/没/就/才/还"（表判断/否定的常见前导，非强调）。
    V55 修正主系表"X 是 Y 的 Z"型误报（V51 仍命中"是年轻人生活中的重要部分"等）：
    - **核心收紧**："的"前一个词的词性必须是【动词】v/vn/vd/vi/vg；
    - **强调句的状语特征**："是"后第一个实词必须是状语 p/nd/nt/d（介词/时间/方位/副词），
      即典型模式"是+状语+动词+的"（"是**昨天来**的""是**在北京出生**的"）。
      而"X是Y的Z"主系表的 Y 通常以 r/n/nd 开头（"是**我**最喜欢的事""是**周末生活**的一部分"）
      或以 v 开头作主谓（"是**住学校**的宿舍"），不属强调句；
    - **特指疑问检测**：若"的"前含疑问代词（r 词性 + "哪儿/哪里/什么/谁/几/怎么"等）
      且"的"前是 v，整段视为特指疑问 → 判 False。
    """
    s, e = match.span()
    seg = match.group(0)
    if "是" not in seg or "的" not in seg:
        return False
    in_seg = [(w.strip(), p) for w, p in zip(words, poses) if w.strip()]
    if not in_seg:
        return False
    shi_idx = [i for i, (w, _p) in enumerate(in_seg) if w == "是"]
    if not shi_idx:
        return False
    i_shi = shi_idx[-1]  # 就近：最后一个"是"
    # 排除前导判断/否定词（"不是/没是/就是/才是/还是"）
    if i_shi > 0 and in_seg[i_shi - 1][0] in ("不", "没", "就", "才", "还"):
        return False
    de_idx = next((j for j in range(i_shi + 1, len(in_seg))
                   if in_seg[j][0].endswith("的")), None)
    if de_idx is None:
        return False
    window = [p for _w, p in in_seg[i_shi + 1: de_idx]]
    if len(window) > 6:  # 窗口过长 → 非强调（判断/叙述句混入）
        return False
    if not window:
        return False
    # V55-1："的"前一个词必须是动词（"是昨天**来**的""是在北京**出生**的"）。
    de_prev_pos = in_seg[de_idx - 1][1] if de_idx - 1 >= i_shi + 1 else ""
    if de_prev_pos not in ("v", "vn", "vd", "vi", "vg"):
        return False
    # V55-2：强调句的"是"后第一个实词必须是状语 p/nd/nt/d 或 方式动词 v
    # （介词/时间/方位/副词/方式结构）。"是+状语+动词+的"是强调句典型结构；
    # "乘飞机去的/开车来的/走路去的"中 LTP 把"乘/开车/走路"标为 v（方式动词），
    # 若只允许 p/nd/nt/d 会漏掉"强调方式"这一重要子类（V55 遗留召回损失）。
    # 若首词是 r/n（代词/名词）开头 → 主系表或主谓句 → 判 False。
    first_pos = in_seg[i_shi + 1][1]
    if first_pos not in ("p", "nd", "nt", "d", "v"):
        return False
    # V55-2b：定中结构"X是Y的Z"（Z 为数量词短语）排除。
    # "是周末生活【的一部分】"中"的"后紧接数量词"一"（m），整段是定中；
    # 强调句"是在食堂吃的【饭】"中"的"后是名词宾语（n），应保留（V55 曾误伤）。
    # 仅在"的"后是数量词 m/q 时判定中，单名词宾语不排除。
    if de_idx + 1 < len(in_seg):
        next_pos = in_seg[de_idx + 1][1]
        if next_pos in ("m", "q"):
            return False
    # V55-3：特指疑问"是 X 的"——含 r 词性疑问词 → 判 False。
    _interrog_r = {"哪", "哪儿", "哪里", "什么", "谁", "怎么", "怎样", "为什么"}
    has_interrog = any((in_seg[k][0] in _interrog_r and in_seg[k][1] == "r")
                        for k in range(i_shi + 1, de_idx))
    if has_interrog:
        return False
    return True


def analyze_foreign_ltp(text):
    """完全本地：LTP 分块分析（按完整句子切块≤400字，合并字/词/语法类型集合后统一计分）
    与程勇2023公式一致；LTP 不可用时直接抛出 RuntimeError，不静默回退 jieba。"""
    char_levels, word_levels, grammar_rules = _get_gf0025()
    word_pos_levels = _get_word_pos_levels()
    char_counts = [0] * 7
    word_counts = [0] * 7
    gram_counts = [0] * 7
    char_detail = [set() for _ in range(7)]
    word_form_detail = [set() for _ in range(7)]  # 各级别命中的词形（词形去重）
    word_detail = [set() for _ in range(7)]       # 词/词性标签（仅展示）
    word_pos_detail = [set() for _ in range(7)]   # (词形, 词性) 条目（诊断，仅精确词性命中）
    word_pos_fallback_detail = [set() for _ in range(7)]  # 词性不匹配、按词形兜底命中的条目（诊断，含预测词性）
    word_form_fallback_detail = [set() for _ in range(7)]  # 按词形唯一键的兜底词种（防同词多词性重复计数）
    gram_detail = [set() for _ in range(7)]
    per_sentence = []
    words_per_sent = []
    poses_per_sent = []
    beyond_words = {}
    beyond_buckets = {"ordinary": {}, "proper": {}, "number": {}, "fragment": {}, "phrase": {}}
    word_ambiguous = set()                     # (词形,词性)仍无法消歧的多级候选词
    word_ambiguous_detail = {}                 # 词形 -> [候选等级列表]（含词性标签，供人工复核）
    token_audit = []                           # 逐 Token 审计：句子/词形/词性/来源/等级/是否计分
    gram_covered = [set() for _ in range(7)]
    gram_triggered = [set() for _ in range(7)]
    gram_events = []  # 逐条命中证据（V36：句子/编号/类型/状态/匹配片段，供四份 CSV）
    ltp_grammar_hits = [0] * 7                # 各等级 LTP 依存句法检测命中数（可信）
    ltp_grammar_triggered = [set() for _ in range(7)]  # 各等级 LTP 命中编号（含候选）
    ltp_grammar_trusted = [set() for _ in range(7)]    # 各等级 LTP 命中编号（可信）

    # 1) 按完整句子切块：正常句子逐句输入；仅单句超 400 字时按边界/重叠拆分，
    #    避免合并句子改变 LTP 分词上下文，同时防长句被静默截断。
    chunks = _chunk_text(text)

    # 1.1) 建立 块 → 原文本句子序号 映射（token 审计按真实句子追踪；超长句拆出的
    #      多个块归属同一句）。
    chunk_sent_ids = []
    _sent_no = 0
    for _sent in split_sentences(text):
        if not _sent.strip():
            continue
        _sent_no += 1
        if len(_sent) <= 400:
            chunk_sent_ids.append(_sent_no)
        else:
            # 超长句拆出的块数 = _long_sentence_chunks 的块数
            _n = len(_long_sentence_chunks(_sent, 400))
            chunk_sent_ids.extend([_sent_no] * _n)
    if len(chunk_sent_ids) != len(chunks):
        # 极端兜底（如文本为空/切分差异）：退化为块序号
        chunk_sent_ids = list(range(1, len(chunks) + 1))

    # 2) LTP 一次处理所有块；不可用则报错
    ltp_results = _ltp_analyze_chunks(chunks)
    if ltp_results is None:
        raise RuntimeError(
            "LTP 不可用：请安装 LTP 环境，或设置 LTP_PYTHON 指向含 LTP 的 Python 解释器，"
            "或运行 download_ltp_model.py 下载模型。"
        )

    # 2.0) 完整性校验：LTP 必须覆盖全部输入块，不允许静默截断
    _verify_ltp_coverage(chunks, ltp_results)

    # 2.1) 语法点编号索引 + LTP 依存句法检测器（真正接入主流程）
    gpi = _grammar_point_index()
    import grammar_ltp_detectors as gld

    # 3) 合并各块的字/词/语法类型集合与频次
    for ci, res in enumerate(ltp_results):
        chunk = chunks[ci]
        sent_no = chunk_sent_ids[ci] if ci < len(chunk_sent_ids) else ci + 1  # 真实句子序号（超长句拆块同句）
        words = [w.strip() for w in res.get("cws", [])]
        poses = res.get("pos", [])
        for ch in chunk:
            lv = char_levels.get(ch)
            if lv:
                char_counts[lv - 1] += 1
                char_detail[lv - 1].add(ch)
        words_per_sent.append(words)
        poses_per_sent.append(poses)
        sent_gram = [set() for _ in range(7)]
        chunk_gram_nos = set()  # 本句已计分的标准语法点编号（正则与LTP去重）
        for w, p in zip(words, poses):
            if not w:
                continue
            lv, amb, src = _lookup_word_level_detail(word_levels, word_pos_levels, w, p)
            # 逐 Token 审计：记录来源（pos=词性精确 / empty=空词性 / form=词形兜底 /
            # None=未定级）、最终等级、是否计入词种分值与分类
            if lv:
                word_counts[lv - 1] += 1
                word_detail[lv - 1].add(_word_detail_label(w, p))
                if src == "form":
                    word_pos_fallback_detail[lv - 1].add(_word_entry_key(w, p))  # 诊断：含预测词性
                    word_form_fallback_detail[lv - 1].add(w)  # 词种：词形唯一键，防同词多词性重复
                else:
                    word_pos_detail[lv - 1].add(_word_entry_key(w, p))
                # 诊断口径：记录词形（正式词种按 (词形,词性) 统计）
                word_form_detail[lv - 1].add(w)
                if amb:
                    word_ambiguous.add(w)
                    # 收集 (词形,词性) 全部候选等级，输出供人工复核（多级同形同词性）
                    cands = []
                    for cat in (_pos_to_word_cats(p) or []):
                        for l in word_pos_levels.get((w, cat), []):
                            if l not in cands:
                                cands.append(l)
                    if cands:
                        word_ambiguous_detail.setdefault(w, [])
                        for l in cands:
                            if l not in word_ambiguous_detail[w]:
                                word_ambiguous_detail[w].append(l)
                token_audit.append({
                    "sentence": sent_no, "word": w, "pos": p,
                    "source": src, "level": lv, "ambiguous": amb,
                    "scored": True,
                    "category": "in_vocab",
                })
            elif re.fullmatch(r'[\u4e00-\u9fff]+', w) and p not in ("w", "wp"):
                beyond_words.setdefault(w, p)
                _add_beyond_word(beyond_buckets, w, p, word_levels)
                token_audit.append({
                    "sentence": sent_no, "word": w, "pos": p,
                    "source": None, "level": 0, "ambiguous": False,
                    "scored": False,
                    "category": _classify_beyond_word(w, p, word_levels),
                })
            else:
                token_audit.append({
                    "sentence": sent_no, "word": w, "pos": p,
                    "source": None, "level": 0, "ambiguous": False,
                    "scored": False,
                    "category": "punctuation" if (p or "").startswith("w") else "out_of_vocab",
                })
        pos_index, sent_rules = _get_grammar_index()
        for w, p in zip(words, poses):
            # V55：助词(u)规则消歧。同一 u 词可能匹配多条助词规则（20/21/22/79/80/81/82/380 等），
            # 若全部记入会导致同一"的"算作 20+82 两种（跨级别双计），同一"了"算作 21+22 两种（同级别双计）。
            # 解决：检查 _UHLP_DISAMBIG，若当前 u 词有首选 no，只接受命中的 no 等于首选 no 的那条规则。
            _u_pref = _UHLP_DISAMBIG.get(w) if p == "u" else None
            for creg, lv, name, no, trusted in pos_index.get(p, ()):
                try:
                    if not creg.fullmatch(w):
                        continue
                    if _u_pref is not None and no != _u_pref:
                        # 同字面被多条助词规则命中，按 _UHLP_DISAMBIG 优先级只保留首选 no
                        continue
                    if no and trusted:
                        gram_counts[lv - 1] += 1
                        chunk_gram_nos.add(no)
                        gram_detail[lv - 1].add(name)
                        sent_gram[lv - 1].add(name)
                    if no:
                        gram_triggered[lv - 1].add(no)
                        gram_events.append({
                            "sentence_id": sent_no, "chunk_id": ci,
                            "level": lv, "no": no, "name": name,
                            "detector_type": "regex",
                            "status": "trusted" if trusted else "candidate",
                            "matched_span": w[:120],
                        })
                        if trusted:

                            gram_covered[lv - 1].add(no)
                except re.error:
                    pass
        for creg, lv, name, no, pos_window, trusted in sent_rules:
            try:
                m = creg.search(chunk)
                if not m:
                    continue
                # V46 词性窗口版 search：pos_window 为字符串词性代码（'v'/'a'）时，
                # 要求匹配片段覆盖的 token 中至少一个词性匹配；True 保持"是……的"旧逻辑
                if isinstance(pos_window, (str, tuple)):
                    if not _match_token_pos_in_span(m, words, poses, pos_window):
                        continue
                elif pos_window and not _is_shi_de_emphasis(m, words, poses):
                    continue
                if _is_frozen_quad(m.group(0)):
                    continue
                if no and trusted:
                    gram_counts[lv - 1] += 1
                    chunk_gram_nos.add(no)
                    gram_detail[lv - 1].add(name)
                    sent_gram[lv - 1].add(name)
                if no:
                    gram_triggered[lv - 1].add(no)
                    gram_events.append({
                        "sentence_id": sent_no, "chunk_id": ci,
                        "level": lv, "no": no, "name": name,
                        "detector_type": "regex",
                        "status": "trusted" if trusted else "candidate",
                        "matched_span": m.group(0)[:120],
                    })
                    if trusted:

                        gram_covered[lv - 1].add(no)
            except re.error:
                pass

        # 3.1) LTP 依存句法检测器（真正进入主流程：分词+词性+依存）
        dep = res.get("dep") or {}
        ltp_hits = set()
        if dep:
            try:
                tokens = gld.parse_sentence(words, poses, dep)
                ltp_hits = gld.detect_triggered(tokens)
            except Exception:
                ltp_hits = set()
        for no in sorted(ltp_hits, key=lambda x: int(x) if x.isdigit() else 10**9):
            info = gpi.get(no)
            if not info:
                continue
            lv = info["level"]
            name = info["name"]
            ltp_grammar_triggered[lv - 1].add(no)
            gram_triggered[lv - 1].add(no)
            trusted_ltp = gld.is_trusted(no)
            gram_events.append({
                "sentence_id": sent_no, "chunk_id": ci,
                "level": lv, "no": no, "name": name,
                "detector_type": "ltp_dep",
                "status": "trusted" if trusted_ltp else "candidate",
                "matched_span": "",
            })
            if not trusted_ltp:
                continue  # 候选特征：只作候选提示，不进分级明细/逐句语法
            ltp_grammar_trusted[lv - 1].add(no)
            gram_covered[lv - 1].add(no)
            gram_detail[lv - 1].add(name)
            sent_gram[lv - 1].add(name)
            if no not in chunk_gram_nos:
                gram_counts[lv - 1] += 1
                ltp_grammar_hits[lv - 1] += 1
                chunk_gram_nos.add(no)
        if chunk.strip():
            per_sentence.append({"sentence": chunk.strip(),
                                 "grammar": [sorted(s) for s in sent_gram]})

    # word_form_detail 仅作诊断：按词形去重，保留最低实际命中等级
    _min_observed = {}
    for _lv, _s in enumerate(word_form_detail, 1):
        for _w in _s:
            if _w not in _min_observed or _lv < _min_observed[_w]:
                _min_observed[_w] = _lv
    word_form_detail = [set() for _ in range(7)]
    for _w, _lv in _min_observed.items():
        word_form_detail[_lv - 1].add(_w)

    # 4) 程勇2023公式：D=(1/3)Σ n*(r_char(n)+r_word(n)+r_grammar(n))
    # r 为各级“种类数”占全文各级总种类数的比例（字种、词种、语法点类型种数）。
    # 词种字段：精确命中 (词形,词性) 与词性兜底 (src=form) 合并统计——
    # 兜底词已按词形定级，不应静默排除出分值；另单独输出兜底词种数供诊断。
    # 统一词形键：若某词形已有【精确命中】(src=pos/empty)，则其兜底词性不再额外计词型，
    # 避免同一词同时输出“很/副”与“很”算 2 个词种。
    # 元素统一为 (词形, 词性) 元组：精确命中 (w, cat)；独立兜底用 (w, "")，
    # 保证 word_unique_counts / 全局并集 / stats.unique_words 三者口径一致。
    # 跨等级去重：只要某词形在【任一等级】精确命中过，其余等级的兜底词形
    # 均不再额外计词型（如“报/动”七级精确 + 三级词形兜底 → 只算七级精确）。
    _all_exact_forms = {w for s in word_pos_detail for (w, _c) in s}
    word_effective_detail = []
    for i in range(7):
        eff = set(word_pos_detail[i])
        for w in word_form_fallback_detail[i]:
            if w not in _all_exact_forms:
                eff.add((w, ""))
        word_effective_detail.append(eff)
    char_types = [len(s) for s in char_detail]
    word_types = [len(s) for s in word_effective_detail]  # 词种按 (词形,词性) + 独立词形兜底
    gram_types = [len(s) for s in gram_covered]
    tc, tw, tg = sum(char_types), sum(word_types), sum(gram_types)
    r_char = [c / tc if tc else 0 for c in char_types]
    r_word = [c / tw if tw else 0 for c in word_types]
    r_gram = [c / tg if tg else 0 for c in gram_types]
    score = sum((i + 1) * (r_char[i] + r_word[i] + r_gram[i]) for i in range(7)) / 3.0

    # stats.unique_words 用 【全局并集】口径（含 (词形,词性) 精确 + (词形,"") 兜底），
    # 与 word_unique_counts（逐级求和）自然不同（跨级词形双计），但两者内部自洽；
    # 各自不再“随机漂移”。
    _stats = _site_stats(text, words_per_sent, word_levels, poses_per_sent, word_pos_levels)
    _stats["unique_words"] = len(set().union(*[set(s) for s in word_effective_detail])) if word_effective_detail else 0

    return {
        "country": "foreign",
        "score": round(score, 4),
        "level": gf0025_level(score),
        "age": "",
        "stats": _stats,
        "char_counts": char_counts,
        "word_counts": word_counts,
        "grammar_counts": gram_counts,
        "char_detail": [sorted(s) for s in char_detail],
        "beyond_words": beyond_words,  # P2-2 超纲词/专名 {词: 词性}（含专名/数字/单字残片，仅诊断）
        "beyond_words_clean": beyond_buckets["ordinary"],  # 普通超纲词，可作人工复核候选
        "beyond_proper_nouns": beyond_buckets["proper"],   # 专名/缩略词
        "beyond_number_like": beyond_buckets["number"],    # 数字组合/数量表达式
        "beyond_fragments": beyond_buckets["fragment"],    # 单字未登录/分词残片
        "beyond_phrases": beyond_buckets["phrase"],        # 短语/分词结果
        "word_ambiguous": sorted(word_ambiguous),  # (词形,词性)无法消歧的多级候选词
        "word_ambiguous_detail": {k: sorted(v) for k, v in word_ambiguous_detail.items()},  # 词形→候选等级（人工复核）
        "word_form_detail": [sorted(s) for s in word_form_detail],  # 各级别词形（词形去重）
        "word_pos_detail": [sorted(_word_detail_label(w, cat) for (w, cat) in s) for s in word_pos_detail],  # 精确命中
        # 兜底词单独计数/展示：词性不匹配但按词形定级的词种（词形唯一键）
        # 若某词形已有精确命中（如“很/副”），其兜底词形不再额外计词型
        "word_pos_fallback_detail": [sorted(_word_detail_label(w, cat) for (w, cat) in s) for s in word_pos_fallback_detail],
        "word_fallback_unique_counts": [
            len([w for w in word_form_fallback_detail[i] if w not in _all_exact_forms])
            for i in range(7)
        ],
        "word_fallback_detail": [
            sorted(w for w in word_form_fallback_detail[i] if w not in _all_exact_forms)
            for i in range(7)
        ],
        # C-修复：语法点覆盖统计（按级别）
        "grammar_catalog": _grammar_catalog(),  # {lv: [(no, name, det)]} 官方全量
        "grammar_covered": [sorted(s) for s in gram_covered],  # 各级命中的标准编号（可信覆盖）
        "grammar_triggered": [sorted(s) for s in gram_triggered],  # 候选触发编号
        "grammar_events": gram_events,  # 逐条命中证据（V36：sentence_id/no/name/level/类型/状态/匹配片段）
        "ltp_grammar_hits": ltp_grammar_hits,  # 各级 LTP 依存句法检测命中数（可信）
        "ltp_grammar_triggered": [sorted(s) for s in ltp_grammar_triggered],  # 各级 LTP 命中编号（含候选）
        "ltp_grammar_trusted": [sorted(s) for s in ltp_grammar_trusted],    # 各级 LTP 命中编号（可信）
        # 正式展示口径：精确词 + 兜底词（均计入词种/分值；与 word_unique_counts 统一键）
        "word_detail": [sorted(_format_effective_detail(word_pos_detail[i], word_form_fallback_detail[i], _all_exact_forms)) for i in range(7)],
        "grammar_detail": [sorted(s) for s in gram_detail],
        "per_sentence": per_sentence,
        "token_audit": token_audit,  # 逐 Token 审计（句子/词形/词性/来源/等级/是否计分/分类）
        "beyond_chars": sorted(set(c for c in text if "一" <= c <= "鿿") - set(char_levels.keys())),
        "char_unique_counts": [len(s) for s in char_detail],
        "word_unique_counts": [len(s) for s in word_effective_detail],  # 词种：精确(词形,词性) + 独立词形兜底（已去重）
        "note": "完全本地离线（LTP 分词）；词表为《国际中文教育中文水平等级标准》(GF0025/HSK 3.0) 官方词汇表，词种数只统计词表内命中词；分值采用论文 GF0025 公式（按字种/词种/语法点类型种数比例计算）",
    }
# ---------------------------------------------------------------------------
_CILIN = {
    "human": {"nr", "nrfg", "nrt"},
    "object": {"n", "ns", "nt", "nz", "nl", "ng", "n"},
    "time_space": {"t", "f"},
    "abstract": {"a", "b", "i", "j", "z"},
    "feature": {"a", "b", "d"},
    "action": {"v", "vn", "vd", "vg", "vi"},
    "mental": {"v", "a"},
    "activity": {"v", "vn"},
    "phenomenon_state": {"a", "v"},
    "relation": {"c", "p"},
    "auxiliary": {"u", "ul", "ug", "ud", "uj", "uz"},
    "respectful": set(),
}


def analyze_native(text):
    """母语汉语模式：程勇母语可读性公式（复刻 readability.py 的 chengyong_readability）

    分词/词性统一使用 LTP（全项目只 LTP；不再用 jieba）。
    """
    char_freq = load_char_freq()
    sents = split_sentences(text)
    if not sents:
        return {
            "country": "native",
            "score": None,
            "level": "—",
            "age": "",
            "note": "母语可读性指数（程勇2020公式），完全本地计算",
            "stats": local_stats(text),
            "semantic_richness": 0,
            "char_detail": [[] for _ in range(7)],
            "word_detail": [[] for _ in range(7)],
            "word_unique_counts": [0] * 7,
            "grammar_detail": [[] for _ in range(7)],
            "per_sentence": [],
            "beyond_chars": [],
            "beyond_words": {},
            "word_ambiguous": [],
            "grammar_catalog": {},
            "grammar_covered": [],
        }

    sentence_lengths = [len(s) for s in sents]
    variance = _variance(sentence_lengths)

    words, poses = [], []
    total_freq = 0.0
    freq_count = 0
    # LTP 分词/词性（与对外汉语生产一致；LTP 不可用时抛错，不静默回退 jieba）
    ltp_results = _ltp_analyze_chunks(sents)
    if ltp_results is None:
        raise RuntimeError(
            "LTP 不可用：请安装 LTP 环境，或设置 LTP_PYTHON 指向含 LTP 的 Python 解释器，"
            "或运行 download_ltp_model.py 下载模型。"
        )
    for res in ltp_results:
        words.extend(w.strip() for w in res.get("cws", []))
        poses.extend(res.get("pos", []))
    for ch in text:
        if ch in char_freq:
            total_freq += char_freq[ch]
            freq_count += 1
    avg_char_freq = total_freq / freq_count if freq_count else 0

    total_words = len(words) or 1
    cat_counts = Counter()
    for pos in poses:
        for cat, pos_set in _CILIN.items():
            if pos in pos_set:
                cat_counts[cat] += 1
                break
    ratios = {cat: cat_counts[cat] / total_words for cat in _CILIN}
    # 语义丰富度 = 出现语义类数 / 【非空】语义类数。
    # V26：原分母用 len(_CILIN)=12，包含"respectful"空类（set()），使丰富度上限被压到
    # 11/12；且 _CILIN 各词性集合存在交集（'a' 属 abstract/feature/mental/phenomenon_state，
    # 'v' 属 action/mental/activity/phenomenon_state），按注册顺序首个匹配 break 使
    # mental/activity/phenomenon_state 恒为 0（词面级分类粒度限制，暂不改类定义，见审查清单）。
    _semantic_classes = [cat for cat, pos_set in _CILIN.items() if pos_set]
    semantic_richness = sum(1 for cat in _CILIN if cat_counts[cat] > 0) / len(_semantic_classes)

    score = (38.36 - 45.65 * avg_char_freq / 1_000_000
             + 54.92 * ratios["relation"]
             - 8.96 * ratios["object"]
             + 11.13 * semantic_richness
             - 12.34 * ratios["action"]
             + 0.012 * variance
             + 20 * ratios["relation"])
    return {
        "country": "native",
        "score": round(score, 4),
        "level": "—",
        "age": "",
        "note": "母语可读性指数（程勇2020公式），完全本地计算；学段映射精度有限（实测初中文本被判小学低年级）",
        "stats": local_stats(text),
        "semantic_richness": round(semantic_richness, 4),
        "char_detail": [[] for _ in range(7)],
        "word_detail": [[] for _ in range(7)],
        "word_unique_counts": [0] * 7,
        "grammar_detail": [[] for _ in range(7)],
        "per_sentence": [],
        "beyond_chars": [],
        "beyond_words": {},
        "word_ambiguous": [],
        "grammar_catalog": {},
        "grammar_covered": [],
    }


def _variance(values):
    if not values:
        return 0
    mean = sum(values) / len(values)
    return sum((v - mean) ** 2 for v in values) / len(values)


# ---------------------------------------------------------------------------
# 统一入口
# ---------------------------------------------------------------------------
def analyze(text, country="foreign"):
    text = text.strip()
    if not text:
        raise ValueError("输入为空")
    if country == "native":
        result = analyze_native(text)
    else:
        # 统一口径：对外汉语一律走生产级 LTP 流程（与 Web/批量报告一致），
        # 避免 Jieba 与 LTP 两套分值/词种口径漂移。LTP 不可用时抛出错误，
        # 不静默回退 Jieba（Jieba 版本保留为 analyze_foreign() 显式快速路径）。
        result = analyze_foreign_ltp(text)
    result["country"] = country
    result["age"] = result.get("age", "")
    return result



if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    demo = "我昨天去商店买东西。商店里有很多人。我买了一件衣服和一双鞋。"
    r = analyze(demo, "foreign")
    print("对外汉语分值:", r["score"], "级别:", r["level"])
    print("统计:", r["stats"])
    for i, name in enumerate(LEVEL_NAMES):
        if r["char_detail"][i] or r["word_detail"][i] or r["grammar_detail"][i]:
            print(f"{name}字({len(r['char_detail'][i])}): {''.join(r['char_detail'][i][:30])}")
            print(f"{name}词({len(r['word_detail'][i])}): {'、'.join(r['word_detail'][i][:30])}")
            print(f"{name}语法({len(r['grammar_detail'][i])}): {'、'.join(r['grammar_detail'][i][:20])}")
    r2 = analyze(demo, "native")
    print("母语可读性:", r2["score"], r2["note"][:30])
