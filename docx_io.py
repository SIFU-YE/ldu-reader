# -*- coding: utf-8 -*-
"""
docx 读件/切分公共库（V32 起替代 batch_ldu.py）

仅包含【读取与切分】层函数，供两类调用方使用：
  1. 生产入口 batch_report_v4.py（输入约定：仅分章文件 x-x-x.docx）；
  2. 论文复现/验证脚本（compare_with_standard.py、reproduce_paper.py、
     ltp_vocab_reproduction.py、grammar_ltp_detectors.py、validate_grammar_rules.py、
     vocabulary_full_reproduction.py）。

不包含任何分析入口（analyze_lesson/CLI main）——分析一律走 core.analyze_foreign_ltp。
"""
import os
import re
import sys
import zipfile

LESSON_TITLE_RE = re.compile(r"^第\s*([0-9一二三四五六七八九十百]+)\s*课\s*(.*)$")
LESSON_TITLE_DOT_RE = re.compile(r"^([0-9]{1,3})\s*[.．、]\s*(.+)$")
# 课标题格式锁定阈值（修订版口径）：
# - 文档中只要出现“第X课”，即锁定“第X课”格式，数字点号格式不再作为课标题。
# - 仅当全文没有任何“第X课”时，才允许使用“1. 标题”格式，且至少需要 2 个点号标题
#   才视为稳定课序列，避免单个正文编号被误判为课标题。
LESSON_TITLE_DOT_MIN_COUNT = 2

# 分章文件：3-2-1.docx；整册文件：3-2.docx（整册仅复现脚本可能使用）
CHAPTER_FILE_RE = re.compile(r"^(\d+-\d+)-(\d+)\.docx$", re.I)
VOLUME_FILE_RE = re.compile(r"^(\d+-\d+)\.docx$", re.I)


def app_dir():
    """程序所在目录：exe 打包后为 exe 所在目录，源码运行时为脚本目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def is_generated_report_file(fname):
    """判断是否为程序生成的报告文件，避免再次扫描分析。"""
    lower = fname.lower()
    return lower.endswith(".docx") and ("_分析报告" in lower or lower == "合并分析报告.docx")


def natural_key(name):
    """自然排序：3-2-10.docx 排在 3-2-2.docx 之后。"""
    return [int(part) if part.isdigit() else part
            for part in re.split(r"(\d+)", name)]


def list_volume_chapter_files(base_dir, prefix):
    """返回某册前缀下的分章文件，按章节号自然排序。"""
    pattern = re.compile(r"^" + re.escape(prefix) + r"-(\d+)\.docx$", re.I)
    found = []
    if not os.path.isdir(base_dir):
        return found
    for fname in os.listdir(base_dir):
        m = pattern.match(fname)
        if m:
            found.append((int(m.group(1)), fname))
    return [fname for _, fname in sorted(found, key=lambda x: x[0])]


def has_volume_chapter_files(base_dir, fname):
    """若目录中同时存在整册文件 x-x.docx 和对应分章 x-x-*.docx，返回 True。"""
    m = VOLUME_FILE_RE.match(fname)
    if not m:
        return False
    return bool(list_volume_chapter_files(base_dir, m.group(1)))


def match_lesson_title(para):
    """识别课标题，支持两种格式：
      1. 第X课 / 第X课 课名（原有）
      2. 1. 课名 / 1．课名（2-5(4).docx 等教材格式）
    返回 (课号, 课名)；非标题返回 None。"""
    s = para.strip()
    m = LESSON_TITLE_RE.match(s)
    if m:
        return m.group(1), (m.group(2).strip() or s)
    m = LESSON_TITLE_DOT_RE.match(s)
    if m:
        name = m.group(2).strip()
        if (name and len(name) <= 30
                and not name.endswith(("。", "！", "？", "；", "："))):
            return m.group(1), name
    return None


def lesson_title_indices(paras):
    """返回课标题段落下标。规则（修订版口径）：
    - 只要文档中存在“第X课”标题，就锁定“第X课”格式，数字点号格式不再作为课标题；
    - 仅当全文没有任何“第X课”时，才允许使用“1. 标题”格式，且点号标题数量
      至少达到 LESSON_TITLE_DOT_MIN_COUNT，才视为稳定课序列。"""
    has_di = any(LESSON_TITLE_RE.match(p.strip()) for p in paras)
    dot_count = 0 if has_di else sum(1 for p in paras if LESSON_TITLE_DOT_RE.match(p.strip()))
    use_dot = (not has_di) and dot_count >= LESSON_TITLE_DOT_MIN_COUNT
    idx = []
    for i, p in enumerate(paras):
        s = p.strip()
        if LESSON_TITLE_RE.match(s):
            idx.append(i)
        elif use_dot and match_lesson_title(s):
            idx.append(i)
    return idx


# 练习节标题特征词：出现在"一、……"式练习小标题中（各册教材练习结构）
_EXERCISE_KEYWORDS = (
    "根据文章", "根据例子", "判断正误", "选词填空", "回答问题",
    "回答以下问题", "讨论下列问题", "选择正确", "排序", "填空",
)


def is_exercise_start(para):
    """判断段落是否为练习起点（正文到此截断，论文口径：统计不包括练习）。

    覆盖两种教材结构：
      1. 显式"练习"标题（第1/3/4/5册）："练习"、"练习一"、"练习"等；
      2. 无"练习"标题、直接以练习小标题开头（第2册）：
         "一、根据文章判断正误。""二、选词填空。""三、根据文章回答问题。"
         "一、根据文章选择正确答案。""一、排序。" 等。
    """
    s = (para or "").strip()
    if not s:
        return False
    # 显式"练习"标题：练习 / 练习：/ 练习一 / 练习（一）/ （练习）等纯标题短段落；
    # 不能把"练习是学习的一部分""练习本"等正文句子误判为练习起点
    if re.match(r"^[（(]?练习[一二三四五六七八九十\d、．.：（(:]*[）)]?$", s) and len(s) <= 8:
        return True
    m = re.match(r"^[一二三四五六七八九十]+、\s*(\S{0,20})$", s)
    if not m:
        return False
    return any(k in m.group(1) for k in _EXERCISE_KEYWORDS)


def extract_docx_paragraphs(path):
    """提取 docx 所有段落文本（含表格外的正文段落）。"""
    import html as _html
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", errors="replace")
    paras = []
    for p in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S):
        line = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p)).strip()
        if line:
            # 反转义 XML 实体（&amp; &lt; &gt; &quot; &apos; 及数字实体），
            # 否则 "A&B <测试>" 会被分析成 "A&amp;B &lt;测试&gt;"
            line = _html.unescape(line)
            paras.append(line)
    return paras


def split_lessons(paras):
    """按课标题切分（第X课 / 1. 标题），返回 [(课号, 课名, 课文正文, 是否含练习)]"""
    title_idx = lesson_title_indices(paras)
    if not title_idx:
        return []
    lessons = []
    for k, start in enumerate(title_idx):
        end = title_idx[k + 1] if k + 1 < len(title_idx) else len(paras)
        title = paras[start]
        m = match_lesson_title(title)
        num, name = m
        # 课文正文：标题后到练习起点之前（论文口径：不包括练习/词语解释）
        body = []
        has_exercise = False
        for p in paras[start + 1:end]:
            if is_exercise_start(p):
                has_exercise = True
                break
            body.append(p)
        lessons.append((num, name, "\n".join(body).strip(), has_exercise))
    return lessons


def iter_lessons_from_docx(path, fname=None):
    """从一个 docx 中产出 (课号, 课名, 正文, 是否含练习)。

    - 整册/普通 docx：按原有“第X课 / 1. 标题”切分；
    - 分章文件 x-x-x.docx：整个文件直接作为一课，课号取文件名中的章节号。
    """
    fname = fname or os.path.basename(path)
    paras = extract_docx_paragraphs(path)
    m = CHAPTER_FILE_RE.match(fname)
    if m:
        text = "\n".join(paras).strip()
        if text:
            yield (m.group(2), os.path.splitext(fname)[0], text, False)
        return
    yield from split_lessons(paras)


def load_volume_lessons(base_dir, prefix):
    """读取一册的全部课文。

    优先读取整册文件 {prefix}.docx 并切分“第X课”；
    若不存在整册文件，则读取 {prefix}-*.docx 分章文件，每个文件作为一课。
    返回 [{"fn","num","title","body","has_exercise","text"}]
    """
    whole = os.path.join(base_dir, f"{prefix}.docx")
    chapter_files = list_volume_chapter_files(base_dir, prefix)
    # 与批处理脚本一致：若同时存在分章文件，优先使用分章，避免整册/分章口径混用
    if not chapter_files and os.path.exists(whole):
        paras = extract_docx_paragraphs(whole)
        out = []
        for num, name, body, has_ex in split_lessons(paras):
            out.append({
                "source": whole,
                "fn": prefix,
                "num": num,
                "title": name,
                "body": body,
                "has_exercise": has_ex,
                "text": body,
            })
        return out

    out = []
    for fname in chapter_files:
        path = os.path.join(base_dir, fname)
        paras = extract_docx_paragraphs(path)
        text = "\n".join(paras).strip()
        m = CHAPTER_FILE_RE.match(fname)
        num = m.group(2) if m else os.path.splitext(fname)[0]
        out.append({
            "source": path,
            "fn": prefix,
            "num": num,
            "title": "",
            "body": text,
            "has_exercise": False,
            "text": text,
        })
    return out
