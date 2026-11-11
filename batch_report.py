#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量 Word 报告 · 完全本地离线版入口

此文件已统一为 batch_report_v4 的兼容入口：
  - 分章文件 x-x-x.docx：生成逐章报告 + 合并 x-x.docx + x-x_分析报告.docx
  - 整册/单个文件：生成单独分析报告
不再包含任何远程调用。
"""

from batch_report_v4 import main

if __name__ == "__main__":
    main()
