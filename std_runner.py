#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""普通版入口（无性能优化）：语法匹配使用全量遍历，与 batch_report_v4 相同功能。"""
import core

core._GRAM_INDEX_USE = False  # 关闭词性索引优化，走全量遍历（普通版行为）

from batch_report_v4 import main  # noqa: E402

if __name__ == "__main__":
    main()
