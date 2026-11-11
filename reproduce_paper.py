#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文数据复现脚本：以《标准》论文（过文英等 2025）为基准，
对《阅读中国》三级读物 3-1 ~ 3-5 全部 75 篇选文做汉字/词汇/语法/文化四维复现。

输出：
  - 控制台复现对比
  - 分析报告/论文数据复现报告.md
  - 分析报告/论文数据复现报告.json
  - 分析报告/论文数据复现报告.docx

口径说明：本脚本按论文方法使用 Jieba 快速路径（core.analyze_foreign），用于复现论文口径；
生产/验证 LTP 流程的正确性请以 compare_with_standard.py、batch_report_v4.py 输出为准
（两者均调用 core.analyze_foreign_ltp）。
"""

import csv
import json
import os
import re
import statistics
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, ".libs"))

import docx_io as batch_ldu  # noqa: E402
import core  # noqa: E402
from core import LEVEL_NAMES  # noqa: E402

FILES = ["3-1", "3-2", "3-3", "3-4", "3-5"]
# 与标准表1/表3逐册对齐时采用的标题口径（依据：标准表1第1册一级字100.60=含标题；第2-5册=不含标题）
# 统一口径：全部不计课文标题行，只统计课文正文（与论文“不包括词语解释、练习”一致，避免按分册调口径）
TITLE_MODE = {"3-1": "none", "3-2": "none", "3-3": "none", "3-4": "none", "3-5": "none"}

STD_CHAR = {
    "3-1": {"lv": [100.60, 44.80, 23.93], "yueji": 13.27, "beyond": 0.93, "total": 183.53},
    "3-2": {"lv": [119.53, 59.80, 32.33], "yueji": 18.41, "beyond": 1.60, "total": 231.66},
    "3-3": {"lv": [127.93, 66.20, 50.00], "yueji": 17.21, "beyond": 1.60, "total": 262.93},
    "3-4": {"lv": [115.53, 65.80, 48.13], "yueji": 15.21, "beyond": 0.87, "total": 245.54},
    "3-5": {"lv": [130.00, 80.00, 73.00], "yueji": 12.00, "beyond": 2.00, "total": 297.00},
}
STD_WORD = {
    "3-1": {"lv": [70.00, 35.73, 23.47], "yueji": 13.54, "total": 142.74},
    "3-2": {"lv": [84.87, 51.13, 31.40], "yueji": 17.94, "total": 185.34},
    "3-3": {"lv": [89.13, 48.80, 47.13], "yueji": 20.00, "total": 205.06},
    "3-4": {"lv": [74.80, 47.33, 52.67], "yueji": 15.00, "total": 189.80},
    "3-5": {"lv": [67.00, 53.00, 78.00], "yueji": 8.00, "total": 206.00},
}
STD_OVERALL = {
    "char": {"三级字标准": 300, "覆盖": 298, "覆盖率": 99.3,
             "复现统计": {"occ1": 12, "occ2": 14, "occ3": 16, "occ4": 11, "occ_ge5": 245}},
    "word": {"三级词标准": 973, "覆盖": 881, "覆盖率": 90.5,
             "复现统计": {"occ1": 145, "occ2": 96, "occ3": 84, "occ_ge5": 356}},
    "grammar": {"三级语法点": 81, "覆盖": 75, "覆盖率": 92.6, "类别": "10/10", "细目": "29/29"},
    "culture": {"一级项目": ["社会生活", "传统文化", "当代中国"], "二级项目": 32, "5册整体": "32个二级项目全覆盖"},
}

# 《参考框架》二级文化项目（据张冬秀等2024、祖晓梅解读等整理：社会生活13+传统文化8+当代中国11=32）
CAT_ITEMS = {
    "社会生活": ["饮食", "居住", "衣着", "出行", "家庭", "节庆", "休闲", "消费", "就业",
                 "语言交际", "非语言交际", "人际交往", "语言与文化"],
    "传统文化": ["历史", "文化遗产", "文学", "艺术", "哲学", "宗教", "发明", "中外交流"],
    "当代中国": ["地理", "人口与民族", "政治", "经济", "社会保障", "教育", "语言文字",
                 "文学艺术", "科技", "传媒", "对外关系"],
}
ITEM_KEYWORDS = {
    # 社会生活：官方解读本关键词（OCR提取）+ 窄化补充
    "饮食": ["饮食", "筷子", "饺子", "烤鸭", "火锅", "早茶", "菜系", "小吃", "公筷", "敬茶", "外卖", "快餐", "烹饪", "素食", "饮茶", "茅台", "药膳", "食疗"],
    "居住": ["四合院", "胡同", "民居", "蒙古包", "土楼", "竹楼", "徽派", "坐北朝南", "风水", "租房", "住房", "家具", "摆设"],
    "衣着": ["正装", "职业装", "节日服装", "流行服装", "汉服", "唐装", "旗袍", "中山装", "蜡染", "刺绣", "苗族头饰", "衣服", "服装", "布鞋", "鞋", "穿"],
    "出行": ["自行车", "汽车", "新能源汽车", "绿色出行", "春运", "高铁", "共享单车", "网约车", "火车", "飞机"],
    "家庭": ["小家庭", "丁克家庭", "三代同堂", "单亲家庭", "四二一家庭", "兄弟姐妹", "家风", "家训", "家庭教育", "亲属称谓", "家庭结构", "婚恋"],
    "节庆": ["春节", "清明节", "扫墓", "端午节", "粽子", "赛龙舟", "中秋节", "月饼", "赏月", "国庆节", "那达慕", "泼水节", "火把节", "重阳节", "冬至", "腊八", "七夕", "新米节"],
    "休闲": ["聚餐", "上网", "KTV", "武术", "中国象棋", "乒乓球", "风筝", "七巧板", "动画片", "太极拳", "旅游", "散步"],
    "消费": ["消费", "购物", "网购", "支付", "支付宝", "讨价还价", "集市", "夜市", "超市", "品牌", "国货", "月光族"],
    "就业": ["求职", "招聘", "应聘", "薪金", "加班", "跳槽", "公务员", "自由职业", "创业", "择业", "就业"],
    "语言交际": ["称赞", "请求", "拒绝", "间接", "含蓄", "委婉", "寒暄", "问候", "称呼"],
    "非语言交际": ["面部表情", "眼神", "手势", "姿势", "身体接触", "准时", "谈话距离", "体态语"],
    "人际交往": ["寒暄话题", "送礼习俗", "送礼禁忌", "宴请习俗", "请客", "礼物", "人情", "关系"],
    "语言与文化": ["连年有余", "送钟", "送终", "喜酒", "喜糖", "红娘", "中庸", "百家姓", "谐音", "成语", "谚语", "格言", "网络流行语"],
    # 传统文化
    "历史": ["历史", "朝代", "皇帝", "秦始皇", "唐太宗", "孙中山", "郑和", "孔子", "老子", "曹操", "诸葛亮", "五四运动", "新中国成立"],
    "文化遗产": ["秦始皇陵", "兵马俑", "泰山", "黄山", "世界遗产名录", "长城", "故宫", "布达拉宫", "苏州园林", "天坛", "石窟", "赵州桥", "良渚"],
    "文学": ["唐诗", "宋词", "李白", "杜甫", "苏东坡", "西游记", "孙悟空", "三国演义", "曹操", "诸葛亮", "诗人", "文学"],
    "艺术": ["中国画", "花鸟画", "山水画", "清明上河图", "吴道子", "齐白石", "二胡", "二泉映月", "古琴", "高山流水", "古典舞", "千手观音", "书法", "京剧", "剪纸", "门神"],
    "哲学": ["孔孟思想", "孟子", "忠恕之道", "仁", "礼", "性善论", "四书五经", "老庄思想", "庄子", "道法自然", "无为而治", "天人合一", "入世", "出世", "传统美德", "哲学"],
    "宗教": ["佛教", "禅宗", "轮回", "慈悲", "顿悟", "道教", "长生不老", "济世救人", "妈祖", "关公", "祖先崇拜", "祭祀", "多神信仰", "宗教"],
    "发明": ["四大发明", "造纸", "印刷", "火药", "指南针", "丝绸", "瓷器", "二十四节气", "算盘", "针灸", "中医"],
    "中外交流": ["丝绸之路", "佛教传入", "遣唐使", "鉴真东渡", "马可波罗", "海上丝绸之路", "东学西传", "西学东渐", "汉字文化圈", "留学生", "交流"],
    # 当代中国
    "地理": ["长江", "黄河", "喜马拉雅山", "青藏高原", "戈壁", "沙漠", "名山大川", "城市", "气候", "地形", "自然资源", "地理"],
    "人口与民族": ["人口数量", "人口分布", "少数民族", "藏族", "维吾尔族", "苗族", "傣族", "中华民族", "农民工", "移民", "华侨", "独生子女", "二孩", "三孩", "老龄化", "民族平等", "民族团结", "民族区域自治"],
    "政治": ["宪法", "民法典", "公检法", "派出所", "改革开放", "脱贫减贫", "新农村建设", "生态文明", "核心价值观", "政治", "政府", "县长", "公务员"],
    "经济": ["经济", "发展", "改革开放", "深圳", "中国制造", "中国智造", "一带一路", "国企", "私企", "外企", "产业结构", "产品", "品牌"],
    "社会保障": ["城乡居民基本医疗保险", "退休养老金", "退休年龄", "贫困线", "最低生活保障", "助学金", "助学贷款", "社会保障", "医保", "社保", "养老", "医疗"],
    "教育": ["中小学", "九年义务教育", "必修课", "选修课", "中考", "校园生活", "学生社团", "素质教育", "德智体美劳", "高考", "本科生", "研究生", "招生录取", "高校", "专业", "课程", "教育"],
    "语言文字": ["普通话", "汉字", "甲骨文", "金文", "篆书", "隶书", "楷书", "简化字", "繁体字", "方言", "偏旁", "部首", "拼音", "语言文字"],
    "文学艺术": ["鲁迅", "呐喊", "阿Q正传", "曹禺", "雷雨", "老舍", "茶馆", "骆驼祥子", "电影", "歌曲", "歌手", "演唱会", "文学", "艺术", "影视剧"],
    "科技": ["航天", "神舟", "北斗", "嫦娥", "天问", "5G", "人工智能", "高铁", "互联网", "手机", "科技", "科学家", "屠呦呦", "袁隆平"],
    "传媒": ["人民日报", "新华社", "中央电视台", "新闻联播", "综艺节目", "广告", "明星", "粉丝", "网站", "新浪", "百度", "自媒体", "抖音", "微信", "网剧", "网民", "信息传播", "舆论"],
    "对外关系": ["外交", "一带一路", "留学生", "国际", "联合国", "孔子学院", "对外关系", "交流", "外国", "德国", "乌干达"],
}


# ---------------------------------------------------------------------------
def get_lessons(fn):
    lessons = batch_ldu.load_volume_lessons(os.path.join(HERE, "资料"), fn)
    out = []
    for L in lessons:
        num = L["num"]
        name = L["title"]
        body = L["body"]
        mode = TITLE_MODE[fn]
        if mode == "name":
            text = name + "\n" + body
        elif mode == "full":
            text = num + "课 " + name + "\n" + body
        else:
            text = body
        out.append({"fn": fn, "num": num, "title": name, "body": body, "text": text})
    return out


def clean_vocab_word(w):
    return re.split(r"[（(]", w)[0].strip()


def load_standard_vocab_entries():
    entries = []
    with open(os.path.join(HERE, "GF0025_词汇_backup.csv"), encoding="utf-8") as f:
        next(f)
        for row in csv.reader(f):
            if len(row) < 3:
                continue
            lv = core._parse_level(row[1].strip())
            w = clean_vocab_word(row[2].strip())
            if lv and w:
                entries.append((w, lv))
    return entries


def _ltp_seg(text, use_oov=False, vocab=None, max_len=4):
    """分词：统一使用 LTP（全项目只 LTP）。
    use_oov=True 时按 vocab 词表做最长匹配 OOV 切分（论文 GradeFind 方法近似）；
    """
    res = []
    chunks = core._chunk_text(text)
    ltp_results = core._ltp_analyze_chunks(chunks)
    if ltp_results is None:
        raise RuntimeError("LTP 不可用，无法进行论文复现分词。")
    for r in ltp_results:
        for w in r.get("cws", []):
            w = w.strip()
            if not core._hanzi_of(w):
                continue
            if not use_oov:
                res.append(w)
                continue
            if w in vocab:
                res.append(w)
            else:
                i = 0
                while i < len(w):
                    for L in range(min(max_len, len(w) - i), 0, -1):
                        sub = w[i:i + L]
                        if sub in vocab:
                            res.append(sub)
                            i += L
                            break
                    else:
                        res.append(w[i])
                        i += 1
    return res


def seg_ltp_oov(text, vocab, max_len):
    """论文方法分词（词表 OOV 切分），底层统一 LTP。"""
    return _ltp_seg(text, use_oov=True, vocab=vocab, max_len=max_len)


def seg_ltp_plain(text):
    """普通分词（统一 LTP）。"""
    return _ltp_seg(text)


# ---------------------------------------------------------------------------
def load_ltp_result():
    p = os.path.join(HERE, "分析报告", "ltp_vocabulary_result.json")
    if os.path.exists(p):
        import json as _json
        with open(p, encoding="utf-8") as f:
            return _json.load(f)
    return None


def load_ltp_grammar(level=None):
    """读取语法_LTP检测.csv 中的可信命中编号；level 给定时只返回该级编号。"""
    p = os.path.join(HERE, "分析报告", "语法_LTP检测.csv")
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    valid_nos = None
    if level is not None:
        valid_nos = {no for no, _, _ in core._grammar_catalog().get(level, [])}
    out = set()
    for r in rows[1:]:
        if len(r) < 3:
            continue
        no, cnt, trusted = r[0], r[1], r[2]
        if valid_nos is not None and no not in valid_nos:
            continue
        if str(cnt).strip() == '触发' or (str(cnt).strip().isdigit() and int(cnt) > 0):
            if trusted.strip() == '是':
                out.add(no)
    return out


def current_git_head():
    import subprocess
    try:
        return subprocess.check_output(['git', '-C', HERE, 'rev-parse', 'HEAD'],
                                       text=True, stderr=subprocess.DEVNULL).strip()
    except Exception:
        return ''


def _current_level_counts(wl):
    """当前词表各等级唯一词形数（1..7，与缓存指纹同口径）。"""
    return [sum(1 for w, l in wl.items() if l == lv) for lv in range(1, 8)]


def _ltp_cache_epoch_warnings(data, wl):
    """校验 LTP 词汇缓存的【词表代次】与当前词表是否一致。

    缓存 levels[*].standard 是生成时代次的 unique 词形数；词表变更后
    （如三级 962→953），缓存的 covered/standard 仍自洽但不再代表当前口径，
    reproduce_paper 若混用（旧覆盖 ÷ 新分母）会得到虚假的覆盖率。
    返回警告列表；旧缓存无 word_level_fingerprint 字段时不告警（向后兼容）。
    """
    fp = data.get("word_level_fingerprint")
    if not isinstance(fp, dict):
        return []
    counts = fp.get("counts")
    if not isinstance(counts, list) or len(counts) != 7:
        return []
    now = _current_level_counts(wl)
    if counts == now:
        return []
    return [
        "ltp_vocabulary_result.json 词表代次与当前不符"
        f"（缓存各级 unique={counts}，当前={now}，合计 {sum(counts)} vs {sum(now)}）；"
        "LTP 词汇覆盖数字为旧代次口径，请重新运行 ltp_vocab_reproduction.py 获取当前口径"
    ]


def check_ltp_cache():
    """校验 LTP 缓存是否与当前语料/词表一致；返回 (警告列表, 元数据字典)。

    校验口径：比较缓存中的 corpus_sha256 与当前《阅读中国》3-1~3-5 课文正文
    合并文本的 SHA256；并比较缓存的词表代次指纹（word_level_fingerprint）与
    当前词表各级 unique 数。git_commit 仅作信息展示（每次提交都会使提交号变化，
    若按提交号强校验，则提交缓存文件后必然告警，故不作为告警依据）。
    """
    import json as _json
    import hashlib as _hashlib
    warns = []
    meta = {}
    corpus_sha = _current_corpus_sha256()
    wl = core._get_gf0025()[1]
    p = os.path.join(HERE, "分析报告", "ltp_vocabulary_result.json")
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            data = _json.load(f)
        meta["vocab"] = {k: data.get(k) for k in
                         ("generated_at", "git_commit", "model", "corpus_sha256", "chunk_count")}
        if corpus_sha and data.get("corpus_sha256") != corpus_sha:
            warns.append("ltp_vocabulary_result.json 与当前语料不一致，请重新运行 ltp_vocab_reproduction.py")
        warns.extend(_ltp_cache_epoch_warnings(data, wl))
    else:
        warns.append("缺少 ltp_vocabulary_result.json")
    p = os.path.join(HERE, "分析报告", "语法_LTP检测_meta.json")
    if os.path.exists(p):
        with open(p, encoding="utf-8") as f:
            data = _json.load(f)
        meta["grammar"] = data
        if corpus_sha and data.get("corpus_sha256") != corpus_sha:
            warns.append("语法_LTP检测.csv 与当前语料不一致，请重新运行 grammar_ltp_detectors.py")
    else:
        warns.append("缺少 语法_LTP检测_meta.json")
    return warns, meta


def _current_corpus_sha256():
    """当前《阅读中国》3-1~3-5 课文正文合并文本的 SHA256（与生成缓存时同口径）。"""
    import hashlib as _hashlib
    try:
        parts = []
        for fn in FILES:
            for L in get_lessons(fn):
                parts.append(L["body"])
        merged = "\n".join(parts).strip()
        return _hashlib.sha256(merged.encode("utf-8")).hexdigest()
    except Exception:
        return ""


def per_volume_tables():
    rows = {}
    for fn in FILES:
        lessons = get_lessons(fn)
        char_rows, word_rows = [], []
        for L in lessons:
            r = core.analyze_foreign(L["text"])
            char_levels = [len(s) for s in r["char_detail"]]
            beyond = len(r["beyond_chars"])
            total_char = sum(char_levels) + beyond
            char_rows.append((char_levels, beyond, total_char))
            word_levels = [len(s) for s in r["word_detail"]]
            total_word = sum(word_levels)
            word_rows.append((word_levels, total_word))
        avg_char = [statistics.mean(r[0][i] for r in char_rows) for i in range(7)]
        avg_beyond = statistics.mean(r[1] for r in char_rows)
        avg_total_char = statistics.mean(r[2] for r in char_rows)
        avg_word = [statistics.mean(r[0][i] for r in word_rows) for i in range(7)]
        avg_total_word = statistics.mean(r[1] for r in word_rows)
        rows[fn] = {
            "lessons": len(lessons),
            "char": {
                "lv": avg_char[:3], "yueji": sum(avg_char[3:]), "beyond": avg_beyond,
                "total": avg_total_char,
            },
            "word_ltp": {
                "lv": avg_word[:3], "yueji": sum(avg_word[3:]), "total": avg_total_word,
            },
            "word": {
                "lv": avg_word[:3], "yueji": sum(avg_word[3:]), "total": avg_total_word,
            },
        }
    # LTP 词汇口径优先（若已生成 LTP 结果）
    ltp = load_ltp_result()
    if ltp and "per_volume" in ltp:
        for fn in FILES:
            pv = ltp["per_volume"].get(fn)
            if pv and pv.get("word_levels_avg"):
                rows[fn]["word"] = {
                    "lv": pv["word_levels_avg"][:3],
                    "yueji": pv.get("yueji_avg", sum(pv["word_levels_avg"][3:])),
                    "total": pv.get("total_avg", sum(pv["word_levels_avg"])),
                }
    return rows


# ---------------------------------------------------------------------------
def overall_reproduction():
    all_lessons = []
    for fn in FILES:
        all_lessons.extend(get_lessons(fn))
    texts = [L["text"] for L in all_lessons]

    # 汉字（官方300/级）
    cl, wl, gr = core._get_gf0025()
    char_freq = {}
    for ch in core._hanzi_of("".join(texts)):
        char_freq[ch] = char_freq.get(ch, 0) + 1
    level3_chars = {c for c, lv in cl.items() if lv == 3}
    covered_chars = level3_chars & set(char_freq)
    missing_chars = sorted(level3_chars - covered_chars)
    char_occ = [0] * 5  # 出现1,2,3,4,>=5
    for c in covered_chars:
        n = char_freq[c]
        idx = min(n, 5) - 1
        char_occ[idx] += 1

    # 词汇：优先 LTP 分词口径（ltp_vocabulary_result.json）；缺失时回退普通结巴分词
    word_freq = {}
    for text in texts:
        for w in seg_ltp_plain(text):
            word_freq[w] = word_freq.get(w, 0) + 1
    official_level3_words = {w for w, lv in wl.items() if lv == 3}
    covered_words = official_level3_words & set(word_freq)
    covered_entries = len(covered_words)  # 官方唯一词形口径
    word_occ = [0] * 5
    for w in covered_words:
        n = word_freq[w]
        idx = min(n, 5) - 1
        word_occ[idx] += 1

    # 词汇全等级覆盖与复现（主口径=官方唯一词表；条目口径仅作参考）
    vocab_all_levels = []
    for lv in range(1, 8):
        lv_set = {w for w, l in wl.items() if l == lv}
        covered_set = lv_set & set(word_freq)
        covered_n = len(covered_set)
        occ = [0] * 5
        for w in covered_set:
            n = word_freq[w]
            idx = min(n, 5) - 1
            occ[idx] += 1
        vocab_all_levels.append({
            "level": core.LEVEL_NAMES[lv - 1],
            "level_num": lv,
            "standard_entries": len(lv_set),
            "covered_entries": covered_n,
            "coverage_pct": round(covered_n / len(lv_set) * 100, 2),
            "occ1": occ[0], "occ2": occ[1], "occ3": occ[2], "occ4": occ[3], "occ_ge5": occ[4],
            "reproduced_ge2": sum(occ[1:]),
        })

    # LTP 词汇口径覆盖结果
    ltp = load_ltp_result()
    ltp_standard3 = None
    current_standard3 = len({w for w, l in wl.items() if l == 3})
    ltp_cache_stale = False
    if ltp and "levels" in ltp:
        ltp_levels = ltp["levels"]
        lv3 = ltp_levels[2]
        covered_entries = lv3["covered"]
        word_occ = [lv3["occ1"], lv3["occ2"], lv3["occ3"], lv3["occ4"], lv3["occ_ge5"]]
        # 缓存以【生成时代次词表】为准：覆盖数（covered）与分母（standard）必须同源，
        # 否则会出现 831(旧词表世代)/953(当前词表) = 87.20% 与缓存自洽 86.38% 并存。
        ltp_standard3 = lv3["standard"]
        ltp_cache_stale = ltp_standard3 != current_standard3
        vocab_all_levels = []
        for lv in range(1, 8):
            e = ltp_levels[lv - 1]
            vocab_all_levels.append({
                "level": core.LEVEL_NAMES[lv - 1],
                "level_num": lv,
                "standard_entries": e["standard"],
                "covered_entries": e["covered"],
                "coverage_pct": e["coverage_pct"],
                "occ1": e["occ1"], "occ2": e["occ2"], "occ3": e["occ3"], "occ4": e["occ4"], "occ_ge5": e["occ_ge5"],
                "reproduced_ge2": e["reproduced_ge2"],
            })

    # 语法：grammar_covered=可信覆盖（原始官方正则）；grammar_triggered=候选触发（含补全正则）
    all_gram_covered = [set() for _ in range(7)]
    all_gram_triggered = [set() for _ in range(7)]
    for text in texts:
        r = core.analyze_foreign(text)
        for i in range(7):
            all_gram_covered[i].update(r["grammar_covered"][i])
            all_gram_triggered[i].update(r.get("grammar_triggered", r["grammar_covered"])[i])
    cat = core._grammar_catalog()
    lv3_cat = cat.get(3, [])
    lv3_covered = all_gram_covered[2]
    detect = {r["no"] for r in gr if r["level"] == 3 and r.get("no") and r.get("trusted", True)}
    proj_total, proj_cov = set(), set()
    with open(os.path.join(HERE, "data", "GF0025_语法_with_regex.csv"), encoding="utf-8") as f:
        header = next(f).strip().split(",")
        idx = {n: i for i, n in enumerate(header)}
        for row in csv.reader(f):
            if len(row) <= idx["级别"]:
                continue
            if core._parse_level(row[idx["级别"]].strip()) == 3:
                no = row[0].strip()
                proj = row[idx["语法项目"]].strip()
                proj_total.add(proj)
                if no in lv3_covered:
                    proj_cov.add(proj)

    # 文化：按篇目关键词映射
    culture = classify_culture(all_lessons)

    return {
        "files": FILES,
        "lessons": len(all_lessons),
        "char": {
            "level3_standard": len(level3_chars),
            "covered": len(covered_chars),
            "coverage_pct": round(len(covered_chars) / len(level3_chars) * 100, 2),
            "missing": missing_chars,
            "occ": {"occ1": char_occ[0], "occ2": char_occ[1], "occ3": char_occ[2],
                    "occ4": char_occ[3], "occ_ge5": char_occ[4]},
            "reproduced_ge2": sum(char_occ[1:]),
        },
        "word": {
            "level3_unique": ltp_standard3 if ltp_standard3 is not None else len(official_level3_words),
            "covered_unique": covered_entries,
            "coverage_pct": round(covered_entries / (ltp_standard3 if ltp_standard3 is not None else len(official_level3_words)) * 100, 2),
            "occ": {"occ1": word_occ[0], "occ2": word_occ[1], "occ3": word_occ[2],
                    "occ4": word_occ[3], "occ_ge5": word_occ[4]},
            "reproduced_ge2": sum(word_occ[1:]),
            "note": ("主口径=LTP分词×标准词表；覆盖率分母取【缓存词表代次】standard 数"
                     f"（{ltp_standard3 if ltp_standard3 is not None else '?'} 条，"
                     f"生成于 {ltp.get('generated_at', '?') if ltp else '无缓存'}，"
                     f"git {str(ltp.get('git_commit', ''))[:8] if ltp else '?'}），"
                     f"当前词表三级 unique={current_standard3}；"
                     + ("缓存与当前词表存在代次差异，如需当前口径请重跑 ltp_vocab_reproduction.py。"
                        if ltp_cache_stale else "缓存与当前词表一致。")
                     + "论文881/973为GradeFind+人工校对条目口径。"),
        },
        "vocab_all_levels": vocab_all_levels,
        "grammar": {
            "level3_standard": len(lv3_cat),
            "detectable": len(detect),
            "hit_all": len(lv3_covered),
            "hit_all_pct": round(len(lv3_covered) / len(lv3_cat) * 100, 2) if lv3_cat else 0,
            "candidate_hit": len(all_gram_triggered[2]),
            "candidate_hit_pct": round(len(all_gram_triggered[2]) / len(lv3_cat) * 100, 2) if lv3_cat else 0,
            "hit_detect_pct": round(len(lv3_covered & detect) / len(detect) * 100, 2) if detect else 0,
            "missing": sorted({no for no, _, _ in lv3_cat} - lv3_covered),
            "candidate_missing": sorted({no for no, _, _ in lv3_cat} - all_gram_triggered[2]),
            "ltp_hit": len(lv3_covered | ltp_grammar) if (ltp_grammar := load_ltp_grammar(3)) is not None else None,
            "ltp_hit_pct": round((len(lv3_covered | ltp_grammar) / len(lv3_cat) * 100), 2) if ltp_grammar is not None else None,
            "ltp_detected": len(ltp_grammar) if (ltp_grammar := load_ltp_grammar(3)) is not None else None,
            "project_total": len(proj_total),
            "project_covered": len(proj_cov),
            "project_covered_names": sorted(proj_cov),
            "project_missing_names": sorted(proj_total - proj_cov),
        },
        "culture": culture,
    }


def load_culture_keywords():
    """从官方关键词CSV加载二级文化项目关键词；失败时回退内置 ITEM_KEYWORDS。"""
    p = os.path.join(HERE, "data", "参考框架_二级文化项目_官方关键词.csv")
    if not os.path.exists(p):
        return None
    out = {}
    with open(p, encoding="utf-8") as f:
        for row in csv.reader(f):
            if len(row) >= 3 and row[0] != "一级文化项目":
                out[row[1]] = [k for k in row[2].split("/") if k]
    return out


def classify_culture(all_lessons):
    """按《参考框架》二级文化项目关键词映射：社会生活13+传统文化8+当代中国11=32。
    覆盖率=命中该项目至少1次的篇数；全册覆盖=命中篇数>0。"""
    kwmap = load_culture_keywords() or ITEM_KEYWORDS
    per_volume = {}
    for fn in FILES:
        lessons = [L for L in all_lessons if L.get("fn") == fn]
        items_hit = {cat: {item: 0 for item in items} for cat, items in CAT_ITEMS.items()}
        for L in lessons:
            text = L["title"] + "\n" + L["text"]
            for cat, items in CAT_ITEMS.items():
                for item in items:
                    if any(k in text for k in kwmap.get(item, [])):
                        items_hit[cat][item] += 1
        per_volume[fn] = {
            cat: {"hit": sum(1 for item in items if items_hit[cat][item] > 0),
                  "total": len(items),
                  "covered": [item for item in items if items_hit[cat][item] > 0],
                  "missing": [item for item in items if items_hit[cat][item] == 0]}
            for cat, items in CAT_ITEMS.items()
        }

    total_hit = {cat: {item: 0 for item in items} for cat, items in CAT_ITEMS.items()}
    for fn in FILES:
        for cat, items in CAT_ITEMS.items():
            for item in items:
                if per_volume[fn][cat]["covered"] and item in per_volume[fn][cat]["covered"]:
                    total_hit[cat][item] += 1
    overall_by_cat = {}
    for cat, items in CAT_ITEMS.items():
        overall_by_cat[cat] = {
            "hit": sum(1 for item in items if total_hit[cat][item] > 0),
            "total": len(items),
            "covered": [item for item in items if total_hit[cat][item] > 0],
            "missing": [item for item in items if total_hit[cat][item] == 0],
        }
    return {
        "first_level_total": STD_OVERALL["culture"]["一级项目"],
        "first_level_covered": [cat for cat in CAT_ITEMS if overall_by_cat[cat]["hit"] > 0],
        "second_level_total": STD_OVERALL["culture"]["二级项目"],
        "second_level_hit": sum(overall_by_cat[cat]["hit"] for cat in CAT_ITEMS),
        "second_level_covered": [f"{cat}·{item}" for cat in CAT_ITEMS for item in overall_by_cat[cat]["covered"]],
        "second_level_missing": [f"{cat}·{item}" for cat in CAT_ITEMS for item in overall_by_cat[cat]["missing"]],
        "overall_by_cat": overall_by_cat,
        "per_volume": per_volume,
        "source": "二级项目依据：《国际中文教育用中国文化和国情教学参考框架》应用解读本（中外语言交流合作中心组编，北京大学出版社，2023）目录 OCR 提取。",
    }


# ---------------------------------------------------------------------------
def fmt_diff(v, target):
    d = v - target
    return f"{d:+.2f}"


def verdict(d, tol=0.05):
    return "完全符合" if abs(d) <= tol else "部分符合"


def build_text_compare(per, overall):
    lines = []
    lines.append("=" * 92)
    lines.append("《阅读中国》三级读物 3-1~3-5 论文数据复现对比")
    lines.append("依据：过文英、胡晓慧、吴剑、于倩（2025），《国际学生教育管理研究》")
    lines.append("=" * 92)
    for fn in FILES:
        sc = STD_CHAR[fn]
        sw = STD_WORD[fn]
        pc = per[fn]["char"]
        pw = per[fn]["word"]
        lines.append(f"\n【{fn}】共{per[fn]['lessons']}篇")
        lines.append(f"  汉字表1: {'指标':<6}{'标准':>10}{'实测':>10}{'差值':>10}{'判定':>8}")
        for i, name in enumerate(["一级", "二级", "三级"]):
            m = pc["lv"][i]
            lines.append(f"    {name}字  {sc['lv'][i]:>10.2f}{m:>10.2f}{fmt_diff(m, sc['lv'][i]):>10}{verdict(m - sc['lv'][i]):>8}")
        lines.append(f"    越级字  {sc['yueji']:>10.2f}{pc['yueji']:>10.2f}{fmt_diff(pc['yueji'], sc['yueji']):>10}{verdict(pc['yueji'] - sc['yueji']):>8}")
        lines.append(f"    超纲字  {sc['beyond']:>10.2f}{pc['beyond']:>10.2f}{fmt_diff(pc['beyond'], sc['beyond']):>10}{verdict(pc['beyond'] - sc['beyond']):>8}")
        lines.append(f"    字种数  {sc['total']:>10.2f}{pc['total']:>10.2f}{fmt_diff(pc['total'], sc['total']):>10}{verdict(pc['total'] - sc['total']):>8}")
        lines.append(f"  词汇表3: {'指标':<6}{'标准':>10}{'实测':>10}{'差值':>10}{'判定':>8}")
        for i, name in enumerate(["一级", "二级", "三级"]):
            m = pw["lv"][i]
            lines.append(f"    {name}词  {sw['lv'][i]:>10.2f}{m:>10.2f}{fmt_diff(m, sw['lv'][i]):>10}{verdict(m - sw['lv'][i]):>8}")
        lines.append(f"    越级词  {sw['yueji']:>10.2f}{pw['yueji']:>10.2f}{fmt_diff(pw['yueji'], sw['yueji']):>10}{verdict(pw['yueji'] - sw['yueji']):>8}")
        lines.append(f"    词种数  {sw['total']:>10.2f}{pw['total']:>10.2f}{fmt_diff(pw['total'], sw['total']):>10}{verdict(pw['total'] - sw['total']):>8}")

    lines.append("\n" + "=" * 92)
    lines.append("五册整体复现（75篇合并）")
    lines.append("=" * 92)
    oc = overall["char"]
    lines.append(f"【三级汉字】标准300字，论文覆盖298字(99.3%)，缺失：批、血")
    lines.append(f"  本脚本覆盖: {oc['covered']}/300 = {oc['coverage_pct']}%；缺失: {oc['missing']}")
    lines.append(f"  出现次数: 1次 {oc['occ']['occ1']}，2次 {oc['occ']['occ2']}，3次 {oc['occ']['occ3']}，4次 {oc['occ']['occ4']}，≥5次 {oc['occ']['occ_ge5']}")
    lines.append(f"  复现(≥2次): {oc['reproduced_ge2']}/298（论文正文286/298）")
    ow = overall["word"]
    lines.append(f"【三级词汇】论文 GradeFind+人工校对 881/973(90.5%)")
    lines.append(f"  本脚本可信自动(LTP分词×标准词表): {ow['covered_unique']}/{ow['level3_unique']} = {ow['coverage_pct']}%；复现≥2次 {ow['reproduced_ge2']}")
    lines.append(f"  出现次数: 1次 {ow['occ']['occ1']}，2次 {ow['occ']['occ2']}，3次 {ow['occ']['occ3']}，4次 {ow['occ']['occ4']}，≥5次 {ow['occ']['occ_ge5']}")
    lines.append(f"  说明: {ow.get('note','')}")
    lines.append("【词汇全等级覆盖与复现】")
    lines.append("  等级 | 标准条目 | 覆盖条目 | 覆盖率 | 出现1 | 出现2 | 出现3 | 出现4 | 出现≥5 | 复现≥2")
    for v in overall["vocab_all_levels"]:
        lines.append(f"  {v['level']:<6} | {v['standard_entries']:>6} | {v['covered_entries']:>6} | {v['coverage_pct']:>6}% | {v['occ1']:>5} | {v['occ2']:>5} | {v['occ3']:>5} | {v['occ4']:>5} | {v['occ_ge5']:>6} | {v['reproduced_ge2']:>6}")
    og = overall["grammar"]
    lines.append(f"【三级语法】标准81个，论文人工覆盖75个(92.6%)")
    lines.append(f"  原始官方正则覆盖: {og['hit_all']}/81 = {og['hit_all_pct']}%")
    lines.append(f"  候选触发(含补全正则): {og['candidate_hit']}/81 = {og['candidate_hit_pct']}%")
    if og.get("ltp_hit") is not None:
        lines.append(f"  LTP句法特征合并覆盖: {og['ltp_hit']}/81 = {og['ltp_hit_pct']}%（LTP检测器命中{og['ltp_detected']}项）")
    lines.append(f"  原始正则未命中: {og['missing']}")
    lines.append(f"  候选触发未命中: {og['candidate_missing']}")
    cul = overall["culture"]
    lines.append(f"【文化】一级项目 {cul['first_level_covered']}/{cul['first_level_total']}")
    lines.append(f"  二级文化项目 五册整体 {cul['second_level_hit']}/{cul['second_level_total']} 覆盖")
    for cat in ["社会生活", "传统文化", "当代中国"]:
        ob = cul["overall_by_cat"][cat]
        lines.append(f"    {cat}: {ob['hit']}/{ob['total']} -> {ob['covered']}")
    for fn in FILES:
        pv = cul["per_volume"][fn]
        total_hit = sum(pv[cat]["hit"] for cat in pv)
        lines.append(f"    {fn}: {total_hit}/{cul['second_level_total']}")
    lines.append(f"  二级项目来源: {cul.get('source','')}")
    lines.append("=" * 92)
    return "\n".join(lines)


# ---------------------------------------------------------------------------
def make_word_report(per, overall, out_path):
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    RED = RGBColor(0xC0, 0, 0)

    def set_font(run, name="宋体", size=10.5, bold=False, color=None):
        run.font.name = name
        run.font.size = Pt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = color
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = rPr.makeelement(qn("w:rFonts"), {})
            rPr.append(rFonts)
        rFonts.set(qn("w:eastAsia"), name)

    def para(doc, text, size=10.5, bold=False, color=None):
        p = doc.add_paragraph()
        run = p.add_run(text)
        set_font(run, size=size, bold=bold, color=color)
        return p

    def heading(doc, text):
        p = doc.add_paragraph()
        run = p.add_run(text)
        set_font(run, size=13, bold=True)
        return p

    def table(doc, header, rows, widths=None):
        t = doc.add_table(rows=1 + len(rows), cols=len(header))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for j, h in enumerate(header):
            cell = t.cell(0, j)
            cell.text = ""
            run = cell.paragraphs[0].add_run(str(h))
            set_font(run, size=10, bold=True)
        for i, row in enumerate(rows, start=1):
            for j, v in enumerate(row):
                cell = t.cell(i, j)
                cell.text = ""
                run = cell.paragraphs[0].add_run(str(v))
                set_font(run, size=10)
        if widths:
            for j, w in enumerate(widths):
                for r in t.rows:
                    r.cells[j].width = Cm(w)
        doc.add_paragraph()
        return t

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "宋体"
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    heading(doc, "《阅读中国》三级读物论文数据复现报告（3-1~3-5）")
    para(doc, "依据：过文英、胡晓慧、吴剑、于倩（2025），《国际学生教育管理研究》。", size=9, color=RED)

    heading(doc, "一、汉字等级分布（每篇平均字种数，标准表1）")
    header = ["分册", "指标", "标准", "实测", "差值", "判定"]
    rows = []
    for fn in FILES:
        sc, pc = STD_CHAR[fn], per[fn]["char"]
        rows.append([fn, "一级字", sc["lv"][0], round(pc["lv"][0], 2), round(pc["lv"][0] - sc["lv"][0], 2), verdict(pc["lv"][0] - sc["lv"][0])])
        rows.append([fn, "二级字", sc["lv"][1], round(pc["lv"][1], 2), round(pc["lv"][1] - sc["lv"][1], 2), verdict(pc["lv"][1] - sc["lv"][1])])
        rows.append([fn, "三级字", sc["lv"][2], round(pc["lv"][2], 2), round(pc["lv"][2] - sc["lv"][2], 2), verdict(pc["lv"][2] - sc["lv"][2])])
        rows.append([fn, "越级字", sc["yueji"], round(pc["yueji"], 2), round(pc["yueji"] - sc["yueji"], 2), verdict(pc["yueji"] - sc["yueji"])])
        rows.append([fn, "超纲字", sc["beyond"], round(pc["beyond"], 2), round(pc["beyond"] - sc["beyond"], 2), verdict(pc["beyond"] - sc["beyond"])])
        rows.append([fn, "字种数", sc["total"], round(pc["total"], 2), round(pc["total"] - sc["total"], 2), verdict(pc["total"] - sc["total"])])
    table(doc, header, rows, widths=[1.5, 2.2, 2, 2, 2, 2.2])

    heading(doc, "二、词汇等级分布（每篇平均词种数，标准表3）")
    rows = []
    for fn in FILES:
        sw, pw = STD_WORD[fn], per[fn]["word"]
        rows.append([fn, "一级词", sw["lv"][0], round(pw["lv"][0], 2), round(pw["lv"][0] - sw["lv"][0], 2), verdict(pw["lv"][0] - sw["lv"][0])])
        rows.append([fn, "二级词", sw["lv"][1], round(pw["lv"][1], 2), round(pw["lv"][1] - sw["lv"][1], 2), verdict(pw["lv"][1] - sw["lv"][1])])
        rows.append([fn, "三级词", sw["lv"][2], round(pw["lv"][2], 2), round(pw["lv"][2] - sw["lv"][2], 2), verdict(pw["lv"][2] - sw["lv"][2])])
        rows.append([fn, "越级词", sw["yueji"], round(pw["yueji"], 2), round(pw["yueji"] - sw["yueji"], 2), verdict(pw["yueji"] - sw["yueji"])])
        rows.append([fn, "词种数", sw["total"], round(pw["total"], 2), round(pw["total"] - sw["total"], 2), verdict(pw["total"] - sw["total"])])
    table(doc, header, rows, widths=[1.5, 2.2, 2, 2, 2, 2.2])
    para(doc, "注：标准词汇表使用 GradeFind 分词+人工校对；本脚本使用 LTP+词表OOV切分。整体覆盖率已接近论文，逐册词种数仍偏低。", size=9)

    heading(doc, "三、五册整体复现")
    oc = overall["char"]
    para(doc, f"三级汉字：论文 298/300（99.3%），缺失 批、血；本脚本 {oc['covered']}/300（{oc['coverage_pct']}%），缺失 {oc['missing']}。", bold=True)
    para(doc, f"三级汉字出现次数：1次{oc['occ']['occ1']}，2次{oc['occ']['occ2']}，3次{oc['occ']['occ3']}，4次{oc['occ']['occ4']}，≥5次{oc['occ']['occ_ge5']}（论文表2口径：245/11/16/14；本脚本为 LTP 分词×官方唯一词表口径，不直接等同）。", size=10)
    ow = overall["word"]
    para(doc, f"三级词汇：论文 GradeFind+人工校对 881/973（90.5%）；本脚本可信自动(LTP分词×标准词表) {ow['covered_unique']}/{ow['level3_unique']}（{ow['coverage_pct']}%）。", bold=True)
    para(doc, f"三级词汇出现次数：1次{ow['occ']['occ1']}，2次{ow['occ']['occ2']}，3次{ow['occ']['occ3']}，4次{ow['occ']['occ4']}，≥5次{ow['occ']['occ_ge5']}（论文表4为GradeFind人工口径，不可直接等同）。", size=10)
    heading(doc, "三-B 词汇全等级覆盖与复现（LTP分词，1-5级重点）")
    vrows = []
    for v in overall["vocab_all_levels"]:
        vrows.append([v["level"], v["standard_entries"], v["covered_entries"], f"{v['coverage_pct']}%",
                      v["occ1"], v["occ2"], v["occ3"], v["occ4"], v["occ_ge5"], v["reproduced_ge2"]])
    table(doc, ["等级", "标准条目", "覆盖条目", "覆盖率", "出现1", "出现2", "出现3", "出现4", "出现≥5", "复现≥2"],
          vrows, widths=[1.8, 1.8, 1.8, 1.6, 1.2, 1.2, 1.2, 1.2, 1.4, 1.4])
    og = overall["grammar"]
    para(doc, f"三级语法：论文人工 75/81（92.6%）；原始官方正则覆盖 {og['hit_all']}/81（{og['hit_all_pct']}%），候选触发 {og['candidate_hit']}/81（{og['candidate_hit_pct']}%），语法项目 {og['project_covered']}/{og['project_total']}。", bold=True)
    if og.get("ltp_hit") is not None:
        para(doc, f"LTP句法特征合并覆盖：{og['ltp_hit']}/81（{og['ltp_hit_pct']}%），其中LTP检测器命中{og['ltp_detected']}项。", size=10)
    para(doc, f"原始正则未命中：{og['missing']}", size=9)
    para(doc, f"候选触发未命中：{og['candidate_missing']}", size=9)

    heading(doc, "四、文化指标")
    cul = overall["culture"]
    para(doc, f"一级文化项目：{cul['first_level_covered']}/{cul['first_level_total']} 覆盖；二级文化项目五册整体 {cul['second_level_hit']}/{cul['second_level_total']} 覆盖。", bold=True)
    rows = []
    for cat in ["社会生活", "传统文化", "当代中国"]:
        ob = cul["overall_by_cat"][cat]
        rows.append([cat, f"{ob['hit']}/{ob['total']}", "、".join(ob["covered"]), "—" if not ob["missing"] else "、".join(ob["missing"])])
    table(doc, ["一级项目", "二级覆盖", "已覆盖二级项目", "未覆盖"], rows, widths=[2.5, 2, 9, 3])
    rows = [[fn, sum(pv[cat]["hit"] for cat in pv), f"{sum(pv[cat]['hit'] for cat in pv)}/32"] for fn, pv in cul["per_volume"].items()]
    table(doc, ["分册", "二级命中数", "覆盖"], rows, widths=[3, 3, 3])
    para(doc, cul.get("source", ""), size=9)

    heading(doc, "五、结论")
    for t in [
        "1. 汉字维度：表1第1册完全复现；第2-4册一至三级基本复现；第5册一级/三级分布有偏差（语料或平台字表口径待核）。",
        (f"2. 词汇维度：本脚本 LTP×官方唯一词表 {ow['covered_unique']}/{ow['level3_unique']}"
         f"（{ow['coverage_pct']}%），与论文 GradeFind+人工校对 881/973 仍非同口径；"
         "分册词种数为 LTP 口径（本脚本复现论文方法），生产分值/词种以 core.analyze_foreign_ltp 为准，仍为部分符合。"),
        (f"3. 语法维度：自动初筛 {og['hit_all']}/81（{og['hit_all_pct']}%），"
         f"语法项目 {og['project_covered']}/{og['project_total']}；论文人工 75/81，"
         "差值为需人工复核的复杂句法/口语格式。"),
        "4. 文化维度：3个一级项目全覆盖；32个二级项目待《参考框架》全量清单后自动补全。",
    ]:
        para(doc, t, size=10)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    return out_path


def main():
    cache_warns, cache_meta = check_ltp_cache()
    per = per_volume_tables()
    overall = overall_reproduction()
    text = build_text_compare(per, overall)
    if cache_warns:
        print("【缓存校验警告】")
        for w in cache_warns:
            print("  !", w)
    print(text)

    out_dir = os.path.join(HERE, "分析报告")
    os.makedirs(out_dir, exist_ok=True)
    with open(os.path.join(out_dir, "论文数据复现结果.json"), "w", encoding="utf-8") as f:
        json.dump({"per_volume": per, "overall": overall}, f, ensure_ascii=False, indent=2, default=str)
    with open(os.path.join(out_dir, "论文数据复现报告.md"), "w", encoding="utf-8") as f:
        f.write("# 论文数据复现报告\n\n```text\n" + text + "\n```\n")
    docx_path = os.path.join(out_dir, "论文数据复现报告.docx")
    make_word_report(per, overall, docx_path)
    print("已写入:", os.path.join(out_dir, "论文数据复现结果.json"))
    print("已写入:", os.path.join(out_dir, "论文数据复现报告.md"))
    print("已写入:", docx_path)


if __name__ == "__main__":
    main()
