#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Excel 指标勾选文件处理工具
执行入口
"""

import os
import sys

# 添加src目录到路径
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from gui import ExcelSelectorGUI
import tkinter as tk

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSelectorGUI(root)
    root.mainloop()
