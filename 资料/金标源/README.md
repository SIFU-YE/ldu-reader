# 金标源数据（V43 校准 · 仅供内部验收）

> ⚠️ 权属与许可声明（重要）：
> - **官方语法_例句.tsv / 官方词汇表.tsv**：抓取自汉考国际 **chinesetest.cn** 官方
>   （仓库 chelsea6502/hsk-2025-data 为同步镜像）。**无开源许可**——仅供本项目
>   **内部验收/校准使用**，**禁止**随 exe 或文档分发给第三方；
> - **hsk_sentences.json**：来自 no7z/hsk-sentences-audio（GitHub/PyPI），
>   数据集 **CC-BY-SA 4.0**（glosses 含 CC-CEDICT 派生；音频 CosyVoice2/Apache-2.0）；
>   引用需遵循署名与相同方式共享；
> - 本目录数据**不参与运行时**计算（运行只读 `data/`），仅被
>   `validate_against_official.py` / `calibrate_levels.py` 等内部校准脚本使用。

## 文件
| 文件 | 内容 | 来源 |
|---|---|---|
| 官方语法_例句.tsv | 593 语法点 / 2055 官方例句（HSK1-7-9） | chinesetest.cn 官方 |
| 官方词汇表.tsv | 官方 HSK 词表（含等级/拼音） | chinesetest.cn 官方 |
| hsk_sentences.json | 4354 句 + 官方等级/语法点标注（1-6 级） | no7z/hsk-sentences-audio (CC-BY-SA 4.0) |

## 校准产物（分析报告/校准/）
- 官方映射表.csv：官方点 → 我们 572 编号（Jaccard ≥0.6，105 条）
- 官方复验_命中率.csv：105 编号 × 官方例句的命中/漏报（总体 76.6%）
- 官方复验_漏报清单.csv：69 条程序漏报的官方例句（规则修正用）
- 官方复验_候选误报Top.csv：例句触发编号 Top（候选噪声粗筛）
- 分级阈值校准.json/.csv：4354 句 官方级 vs 我们级 混淆矩阵 + 阈值扫描
  （当前 19.02% → 最优边界 [1.25,1.6,1.85,2.15,2.4,2.85] 达 29.72%）
