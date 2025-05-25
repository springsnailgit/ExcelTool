#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 表格对比工具 - 简化自适应版 v4.1.1
专注于鼠标拖拽的自适应功能，无额外控制面板
"""

import sys
import os

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 导入主程序
from Excel_Tool import ImprovedExcelCompareGUI
import tkinter as tk

def main_simplified():
    """启动简化版Excel对比工具"""
    
    # 创建主窗口
    root = tk.Tk()
    app = ImprovedExcelCompareGUI(root)
    
    # 设置窗口标题
    root.title("Excel表格对比工具 - 自适应版 v4.1.1")
    
    # 添加提示信息到日志
    app.log("🎯 自适应界面提示：")
    app.log("• 可以拖拽窗口边框任意调整大小")
    app.log("• 比对设置中的两个功能模块会自动等比例调整")
    app.log("• 输出目录区域有最小高度保护，确保始终可操作")
    app.log("• 最小窗口尺寸：900x700，推荐尺寸：1100x850")
    app.log("• 所有组件都会跟随外边框变化进行自适应调整")
    
    # 启动界面
    root.mainloop()

if __name__ == "__main__":
    print("启动Excel表格对比工具 - 简化自适应版 v4.1.1...")
    print("特性：通过鼠标拖拽窗口边框测试自适应效果")
    main_simplified() 