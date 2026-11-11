#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 报告生成：完全本地页面同构输出结构

页面结构 → Word 文档：
  1. 阅读难度分值 / 预测级别 / 适合年龄
  2. 阅读难度分级体系表（分数区间 × 难度等级）
  3. 本文易读性得分及难度等级
  4. 文章基本情况表（总字数/字种/总词数/词种/句子/平均句长/最长句长）
  5. 分级字详细信息表（一级字~七-九级字 + 超纲字）
  6. 分级词详细信息表（一级词~七-九级词）
  7. 分级语法详细信息表（句子 × 各级语法点）
  8. 语法点覆盖统计表
  9. 难度相似文章推荐表
  10. 引用信息（LDU-TGP 版权 + 论文出处）
  11. 方法说明

依赖：python-docx
"""

import os
import sys

# 优先使用仓库内 .libs 离线依赖（python -m pip --target .libs）
_LOCAL_LIBS = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".libs")
if os.path.isdir(_LOCAL_LIBS) and _LOCAL_LIBS not in sys.path:
    sys.path.insert(0, _LOCAL_LIBS)

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

LEVEL_NAMES = ["一级", "二级", "三级", "四级", "五级", "六级", "七-九级"]
SCORE_BOUNDS = ["1-1.4", "1.4-1.75", "1.75-2", "2-2.3", "2.3-2.55", "2.55-2.7", "2.7+"]
RED = RGBColor(0xC0, 0x00, 0x00)


def _set_cn_font(run, name="宋体", size=10.5, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    r = run._element.rPr
    rFonts = r.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = r.makeelement(qn("w:rFonts"), {})
        r.append(rFonts)
    rFonts.set(qn("w:eastAsia"), name)


def _para(doc, text, size=10.5, bold=False, align=None, color=None, space_after=4):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    _set_cn_font(run, size=size, bold=bold, color=color)
    return p


def _heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    _set_cn_font(run, size=14 if level == 1 else 12, bold=True)
    return p


def _table(doc, header, rows, col_widths=None):
    t = doc.add_table(rows=0, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_row = t.add_row()
    for j, h in enumerate(header):
        cell = header_row.cells[j]
        cell.text = ""
        run = cell.paragraphs[0].add_run(str(h))
        _set_cn_font(run, size=10, bold=True)
    for row_data in rows:
        row = t.add_row()
        for j, v in enumerate(row_data):
            cell = row.cells[j]
            cell.text = ""
            run = cell.paragraphs[0].add_run(str(v) if v is not None else "")
            _set_cn_font(run, size=10)
    if col_widths:
        for j, w in enumerate(col_widths):
            for row in t.rows:
                row.cells[j].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def make_report(result, out_path, doc_title="汉语文本阅读难度分析报告"):
    """生成页面同构的 Word 报告。result 来自 core.analyze_foreign / analyze_foreign_ltp。"""
    doc = Document()
    # 页面默认字体
    style = doc.styles["Normal"]
    style.font.name = "宋体"
    style.font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    # 标题区（对应页面 h3.pageTitle）
    _heading(doc, doc_title, 1)
    _para(doc, "LDU-TGP · 鲁东大学文本分级平台", size=10, color=RED, space_after=10)

    # 一、阅读难度评估结果（页面第一块）
    _heading(doc, "一、阅读难度评估结果", 2)
    _para(doc, f"阅读难度分值为：{result['score']}", bold=True, color=RED)
    _para(doc, f"预测文章对应级别为：{result['level']}", bold=True, color=RED)
    age = result.get("age") or "—"
    _para(doc, f"预测适合阅读年龄为：{age}", bold=True, color=RED)
    _para(doc, f"本文易读性得分及难度等级为：{result['score']}，{result['level']}",
          bold=True, color=RED)

    # 二、阅读难度分级体系（页面表格）
    _heading(doc, "二、阅读难度分级体系", 2)
    _table(doc, ["分数区间"] + SCORE_BOUNDS,
           [["难度等级"] + LEVEL_NAMES],
           col_widths=[2.5, 1.6, 1.6, 1.6, 1.6, 1.7, 1.7, 1.7])

    # 三、文章基本情况（页面第二块）
    _heading(doc, "三、文章基本情况", 2)
    s = result.get("stats", {})
    _table(doc, ["指标", "数值", "指标", "数值"], [
        ["总字数", s.get("total_chars", "—"), "字种数", s.get("unique_chars", "—")],
        ["分词Token数（含标点）", s.get("total_words", "—"),
         "纯汉字词数", s.get("total_hanzi_words", "—")],
        ["句子数", s.get("sentences", "—"), "平均句长", s.get("avg_sent_len", "—")],
        ["词种数", s.get("unique_words", "—"), "", ""],
    ], col_widths=[3, 3, 3, 3])

    # 四、分级字详细信息（页面表格，本地补全；增加字种数与占比，与《标准》表1同口径）
    _heading(doc, "四、分级字详细信息", 2)
    rows = []
    char_total = 0
    for i, name in enumerate(LEVEL_NAMES):
        items = result["char_detail"][i] if result.get("char_detail") else []
        char_total += len(items)
    beyond = result.get("beyond_chars", [])
    char_total += len(beyond)
    for i, name in enumerate(LEVEL_NAMES):
        items = result["char_detail"][i] if result.get("char_detail") else []
        pct = f"{len(items) / char_total * 100:.1f}%" if char_total else "—"
        rows.append([f"{name}字", len(items), pct, "、".join(items) if items else "—"])
    yueji_chars = sum(len(result["char_detail"][i]) for i in range(3, 7)) if result.get("char_detail") else 0
    pct = f"{yueji_chars / char_total * 100:.1f}%" if char_total else "—"
    rows.append(["越级字(四-七九级合计)", yueji_chars, pct, "—"])
    pct = f"{len(beyond) / char_total * 100:.1f}%" if char_total else "—"
    rows.append(["超纲字", len(beyond), pct, "、".join(beyond) if beyond else "—"])
    _table(doc, ["级别", "字种数", "占比", "汉字"], rows, col_widths=[2.2, 1.6, 1.6, 11])

    # 五、分级词详细信息（按 (词形,词性) 去重；同一词形不同词性分别计数；
    #     本表为【逐级求和】——跨级词（如“很”）可在多个等级各计 1 次，其和可能
    #     大于上表“词种数”（全局并集口径），两者口径不同均属预期）
    _heading(doc, "五、分级词详细信息", 2)
    rows = []
    word_counts = result.get("word_unique_counts")
    if not word_counts:
        # 兼容旧结果/缺少词种数字段时：由分级词明细反推词种数
        word_counts = [len(items) for items in (result.get("word_detail") or [])]
        word_counts = (word_counts + [0] * (7 - len(word_counts)))[:7]
    word_total = sum(word_counts)
    for i, name in enumerate(LEVEL_NAMES):
        items = result["word_detail"][i] if result.get("word_detail") else []
        n = word_counts[i] if i < len(word_counts) else len(items)
        pct = f"{n / word_total * 100:.1f}%" if word_total else "—"
        rows.append([f"{name}词", n, pct, "、".join(items) if items else "—"])
    yueji_words = sum(word_counts[3:7]) if len(word_counts) >= 7 else 0
    pct = f"{yueji_words / word_total * 100:.1f}%" if word_total else "—"
    rows.append(["越级词(四-七九级合计)", yueji_words, pct, "—"])
    _table(doc, ["级别", "词种数", "占比", "词语"], rows, col_widths=[2.2, 1.6, 1.6, 11])
    fb_total = sum(result.get("word_fallback_unique_counts") or [])
    if fb_total:
        # 词性兜底词：词形在词表内、但 LTP 词性与词表标注不匹配，按词形定级并计入词种/分值。
        # 单列展示便于人工复核词表词性标注。
        _para(doc, f"词性兜底词（词形命中、词性不匹配后按词形定级；共 {fb_total} 词种，已计入上方词种数）：",
              size=10)
        fb_rows = []
        fb_detail = result.get("word_fallback_detail") or [[] for _ in range(7)]
        for i, name in enumerate(LEVEL_NAMES):
            items = fb_detail[i] if i < len(fb_detail) else []
            if items:
                fb_rows.append([name, len(items), "、".join(items)])
        if fb_rows:
            _table(doc, ["级别", "词种数", "词语"], fb_rows, col_widths=[2.2, 1.6, 12])

    # 六、超纲词与专名（P2-2：词表外词单列）
    # 修复口径：普通超纲词与“专名/数字组合/单字残片”分开列示；
    # 只有普通超纲词才作为人工教学词表复核候选，专名/数字/残片仅作诊断。
    _heading(doc, "六、超纲词与专名", 2)

    def _beyond_table(title, data):
        if not data:
            return False
        _para(doc, title, size=10)
        pos_names = {"ns": "地名", "nt": "机构/时间专名", "nr": "人名", "nh": "人名",
                     "ni": "机构名", "nz": "其他专名", "j": "简称/缩略",
                     "n": "名词", "v": "动词", "a": "形容词", "d": "副词",
                     "m": "数词", "q": "量词", "r": "代词"}
        from collections import defaultdict
        bypos = defaultdict(list)
        for w, p in data.items():
            bypos[p].append(w)
        rows = []
        for p, ws in sorted(bypos.items(), key=lambda x: -len(x[1])):
            shown = "、".join(sorted(ws)[:40]) + ("……" if len(ws) > 40 else "")
            rows.append([pos_names.get(p, p), len(ws), shown])
        _table(doc, ["类型", "词数", "词语（节选）"], rows, col_widths=[2.5, 2, 11])
        return True

    bw = result.get("beyond_words") or {}
    bw_clean = result.get("beyond_words_clean")
    if bw_clean is None:
        # 旧结果兼容：没有分类字段时退回原全量清单
        bw_clean = bw
    has_clean = _beyond_table(
        f"以下 {len(bw_clean)} 个普通词不在《国际中文教育中文水平等级标准》词表内"
        "（已剔除专名、数字组合、单字残片；建议人工复核后使用，不计入词种数）：",
        bw_clean)
    if not has_clean and not bw:
        _para(doc, "（未识别到词表外词）", color=RED)

    specials = [
        ("专名/缩略词", "beyond_proper_nouns", "地名、人名、机构名、其他专名等，建议作为注释词处理"),
        ("数字组合/数量表达式", "beyond_number_like", "数字、日期、数量组合等，不作为词汇教学词表"),
        ("短语/分词结果", "beyond_phrases", "可切分为多个标准词的短语或分词结果，不作为独立词条"),
        ("单字未登录/分词残片", "beyond_fragments", "单字未登录词，可能为分词残片，需结合上下文人工复核"),
    ]
    for title, key, note in specials:
        data = result.get(key) or {}
        if data:
            _beyond_table(
                f"{title}（{len(data)} 个，{note}）：",
                data)
    wa = result.get("word_ambiguous") or []
    if wa:
        _para(doc, f"多级候选词（(词形,词性)仍无法消歧，暂按最低级计）：{ '、'.join(wa) }", size=9, color=RED)

    # 七、分级语法详细信息（句子 × 各级语法点，页面结构）
    _heading(doc, "七、分级语法详细信息", 2)
    per = result.get("per_sentence", [])
    if per:
        header = ["句子"] + LEVEL_NAMES
        rows = [[ps["sentence"], *["、".join(g) if g else "" for g in ps["grammar"]]]
                for ps in per]
        _table(doc, header, rows, col_widths=[4, 1.6, 1.6, 1.6, 1.6, 1.6, 1.6, 1.6])
    else:
        _para(doc, "（本文未识别到语法点）", color=RED)

    # 八、语法点覆盖统计（C-修复：按标准编号汇总，双口径）
    _heading(doc, "八、语法点覆盖统计", 2)
    lth = result.get("ltp_grammar_hits") or []
    if lth:
        _para(doc, f"LTP依存句法检测命中句次（可信，句次=检测器在分句上命中的次数）："
                   f"{ '，'.join(f'{LEVEL_NAMES[i]} {lth[i]}' for i in range(7) if lth[i]) }",
              size=9, color=RED)
    catalog_raw = result.get("grammar_catalog") or {}
    # JSON 序列化后字典键会变成字符串；统一转回 int，避免 catalog.get(3) 取不到三级
    catalog = {int(k) if str(k).isdigit() else k: v for k, v in catalog_raw.items()}
    covered = result.get("grammar_covered") or []
    triggered = result.get("grammar_triggered") or []
    # 可检测规则集合（有正则且非后缀的规则，来自 core 索引；V36 并入可信 LTP 检测器）
    try:
        import core as _core
        import grammar_ltp_detectors as _gld
        _, _, _gram = _core._get_gf0025()
        _gpi = _core._grammar_point_index()
        detect_by_lv = {}
        for r in _gram:
            if r["level"] and r.get("no") and r.get("trusted", True):
                detect_by_lv.setdefault(r["level"], set()).add(r.get("no", ""))
        # 正则与 LTP 双路径可信并集，保证“可自动检测”分母 = 实际可评分路径
        for no in list(_gld.DETECTORS) + list(_gld.LEVEL_DETECTORS):
            if _gld.is_trusted(no):
                info = _gpi.get(no)
                if info and info["level"]:
                    detect_by_lv.setdefault(info["level"], set()).add(no)
    except Exception:
        detect_by_lv = {}
    if catalog:
        rows = []
        for lv in range(1, 8):
            items = catalog.get(lv, [])
            if not items:
                continue
            cov = set(covered[lv - 1]) if lv - 1 < len(covered) else set()
            detect = detect_by_lv.get(lv, set())
            hit_all = sum(1 for no, _, _ in items if no in cov)
            hit_det = sum(1 for no, _, _ in items if no in cov and no in detect)
            rows.append([LEVEL_NAMES[lv - 1], len(items), len(detect),
                         f"{hit_all}/{len(items)}",
                         f"{hit_det}/{len(detect)}（{hit_det / len(detect) * 100:.1f}%）"
                         if detect else "—"])
        _table(doc, ["级别", "标准语法点", "可自动检测", "命中/全量", "命中/可检测（自动覆盖率）"],
               rows, col_widths=[2.5, 2.2, 2.2, 2.2, 4.5])
        # 各级命中分析（正则 + LTP可信补充，逐级给出）
        lt_trig = result.get("ltp_grammar_triggered") or []
        lt_trust = result.get("ltp_grammar_trusted") or []
        analysis_rows = []
        analysis_data = []
        for lv in range(1, 8):
            items = catalog.get(lv, [])
            if not items:
                continue
            cov = set(covered[lv - 1]) if lv - 1 < len(covered) else set()
            detect = detect_by_lv.get(lv, set())
            hit_all = sum(1 for no, _, _ in items if no in cov)
            hit_det = sum(1 for no, _, _ in items if no in cov and no in detect)
            trust = set(lt_trust[lv - 1]) if lv - 1 < len(lt_trust) else set()
            hit_ltp_supp = len((cov & trust) - detect)
            det_pct = f"{hit_det / len(detect) * 100:.1f}%" if detect else "—"
            all_pct = f"{hit_all / len(items) * 100:.1f}%" if items else "—"
            analysis_rows.append([
                LEVEL_NAMES[lv - 1], len(items), len(detect), hit_det,
                det_pct, hit_ltp_supp,
                f"{hit_all}/{len(items)}（{all_pct}）",
            ])
            analysis_data.append((lv, len(items), len(detect), hit_det,
                                  det_pct, hit_ltp_supp, hit_all, all_pct))
        _table(doc, ["级别", "标准语法点", "可正则自动检测", "正则命中",
                     "自动覆盖率(分母=可正则检测数)", "LTP可信补充", "全量命中"],
               analysis_rows, col_widths=[1.8, 1.8, 2.0, 1.4, 3.0, 1.6, 2.6])

        # 逐级文字分析（一~七九级同口径）
        for lv, n_items, n_detect, hit_det, det_pct, hit_ltp_supp, hit_all, all_pct in analysis_data:
            _para(doc, f"{LEVEL_NAMES[lv - 1]}标准语法点共 {n_items} 个；"
                       f"其中程序可正则自动检测 {n_detect} 个。"
                       f"本文共命中 {hit_all} 个：正则命中 {hit_det} 个"
                       f"（自动覆盖率 {det_pct}，分母=可正则检测数），"
                       f"LTP依存句法检测器可信补充 {hit_ltp_supp} 个；"
                       f"相对标准全量命中率 {all_pct}。",
                  size=9)
            if lv == 3:
                _para(doc, "三级论文参照：论文为 75/81（92.6%，人工语法分析）。程序仅为自动初筛，"
                           "人工复核时需按标准全量 81 个逐一核对，不能直接以自动结果断言覆盖率。",
                      size=9)

        # 全等级语法点细目表（编号 1-572，逐项列出可检测与命中状态）
        all_rows = []
        for lv in range(1, 8):
            items = catalog.get(lv, [])
            cov = set(covered[lv - 1]) if lv - 1 < len(covered) else set()
            detect = detect_by_lv.get(lv, set())
            for no, name, det in items:
                detectable = "✓" if no in detect else "—"
                hit = "✓" if no in cov else "—"
                all_rows.append([no, LEVEL_NAMES[lv - 1], name, det if det else "—",
                                 detectable, hit])
        all_rows.sort(key=lambda r: int(r[0]))
        _table(doc, ["编号", "级别", "语法点", "细目", "可检测", "命中"],
               all_rows, col_widths=[1.2, 1.2, 5, 4, 1.2, 1.2])

        # LTP 依存句法检测明细（全等级，含候选与可信）
        if lt_trig or lt_trust:
            ltp_rows = []
            for lv in range(1, 8):
                items = catalog.get(lv, [])
                if not items:
                    continue
                trig = set(lt_trig[lv - 1]) if lv - 1 < len(lt_trig) else set()
                trust = set(lt_trust[lv - 1]) if lv - 1 < len(lt_trust) else set()
                if not trig and not trust:
                    continue

                def fmt_nos(nos):
                    parts = []
                    for no in sorted(nos, key=lambda x: int(x)):
                        nm = next((name for num, name, _ in items if num == no), "")
                        parts.append(f"{no} {nm}" if nm else no)
                    return "、".join(parts) if parts else "—"
                ltp_rows.append([LEVEL_NAMES[lv - 1],
                                 fmt_nos(trust),
                                 fmt_nos(trig - trust)])
            if ltp_rows:
                _para(doc, "LTP依存句法检测明细（全等级）：", bold=True, size=9)
                _table(doc, ["级别", "LTP可信命中（编号 语法点）", "LTP候选命中（编号 语法点）"],
                       ltp_rows, col_widths=[2, 6.5, 6.5])

        # 正则检出明细（全等级）：与 LTP 明细结构一致，不依赖是否存在 LTP 命中
        regex_rows = []
        for lv in range(1, 8):
            items = catalog.get(lv, [])
            if not items:
                continue
            cov = set(covered[lv - 1]) if lv - 1 < len(covered) else set()
            trig = set(triggered[lv - 1]) if lv - 1 < len(triggered) else set()
            lt_trig_s = set(lt_trig[lv - 1]) if lv - 1 < len(lt_trig) else set()
            lt_trust_s = set(lt_trust[lv - 1]) if lv - 1 < len(lt_trust) else set()
            regex_trust = cov - lt_trust_s
            regex_cand = (trig - lt_trig_s) - regex_trust
            if not regex_trust and not regex_cand:
                continue

            def fmt_nos(nos):
                parts = []
                for no in sorted(nos, key=lambda x: int(x)):
                    nm = next((name for num, name, _ in items if num == no), "")
                    parts.append(f"{no} {nm}" if nm else no)
                return "、".join(parts) if parts else "—"
            regex_rows.append([LEVEL_NAMES[lv - 1],
                               fmt_nos(regex_trust),
                               fmt_nos(regex_cand)])
        if regex_rows:
            _para(doc, "正则检出明细（全等级）：", bold=True, size=9)
            _table(doc, ["级别", "正则可信命中（编号 语法点）", "正则候选命中（编号 语法点）"],
                   regex_rows, col_widths=[2, 6.5, 6.5])
    else:
        _para(doc, "（无语法点覆盖数据）", color=RED)

    # 九、难度相似文章推荐（本地未部署该功能）
    _heading(doc, "九、难度相似文章推荐", 2)
    _table(doc, ["课文名", "教材出处", "难度分数"],
           [["—", "—", "—"]], col_widths=[6, 6, 3.5])
    _para(doc, "注：本地版暂未部署相似文章推荐数据。", size=9)

    # 十、引用信息（页面底部版权 + 论文）
    _heading(doc, "十、引用信息", 2)
    _para(doc, "如需引用该平台，请标注\"LDU-TGP,鲁东大学文本分级平台\"，参考文献如下：", bold=True)
    _para(doc, "对外汉语：\"基于新标准的汉语二语文本阅读难度分级体系构建及应用\"，程勇等，《世界汉语教学》2023年第1期。", size=9)
    _para(doc, "母语汉语：\"基于语文教材语料库的文本阅读难度分级关键因素分析与易读性公式研究\"，程勇等，《语言文字应用》2020年第1期。", size=9)
    _para(doc, f"分析模式：{'对外汉语' if result.get('country') == 'foreign' else '母语汉语'} | "
               f"生成时间：{__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M')}",
          size=9)

    # 十一、方法说明（P3：口径与局限）
    _heading(doc, "十一、方法说明", 2)
    _para(doc, "1. 总字数为全文非空字符数；字种数为纯汉字（不含数字/字母/标点）。", size=9)
    _para(doc, "2. 词种数按“(词形,词性)”统计，仅统计命中《国际中文教育中文水平等级标准》词表的词；"
               "同一词形不同词性分别计数，例如“白/形”和“白/副”计为 2 个词种；"
               "词性不匹配时按词形最低级兜底，计入词种数与分值，并单列在“词性兜底词明细”中供人工复核；"
               "超纲词按普通词、专名/缩略、数字组合、短语/分词结果、单字残片分别列出（见第六章），不计入词种数。", size=9)
    _para(doc, "3. 语法点识别为程序化规则匹配（LTP 词性 + 正则），高级别四字格/句式命中建议人工复核；"
               "类后缀（-化/-子/-儿/-们 等）不做可信自动判定——其词尾正则仅作候选提示（见正则候选命中表），"
               "不计入可信覆盖与自动覆盖率，需人工复核。", size=9)
    _para(doc, "4. 难度分值为综合易读性预测，非严格分级达标证明；"
               "是否严格符合某级需结合字/词/语法分布人工核验。", size=9)
    if result.get("note"):
        _para(doc, f"5. 运行说明：{result['note']}", size=9)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    return out_path


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import core
    text = "我昨天去商店买东西。商店里有很多人。我买了一件衣服和一双鞋。"
    r = core.analyze_foreign(text)
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "report_demo.docx")
    make_report(r, out)
    print("报告已生成:", out, f"({os.path.getsize(out)}B)")
    print("分值:", r["score"], "级别:", r["level"], "统计:", r["stats"])
