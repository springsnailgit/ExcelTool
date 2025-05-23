#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 表格对比工具 - 改进版
功能：精确对比Excel表格内容，仅比对共同列并高亮差异
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
import os
import sys
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
import threading
from pathlib import Path


class ImprovedExcelCompareEngine:
    """改进版Excel对比引擎"""
    
    def __init__(self):
        self.main_table = None
        self.main_columns = []
        self.key_column_index = 0  # 用于行匹配的关键列索引
        self.main_file_name = ""
        self.excluded_columns = set()  # 用户手动排除的列
        
    def load_main_table(self, file_path):
        """加载主对比表格"""
        try:
            self.main_file_name = os.path.splitext(os.path.basename(file_path))[0]
            
            # 读取Excel文件（支持多sheet，这里取第一个）
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            
            # 使用第一个sheet作为主表格
            first_sheet = sheet_names[0]
            self.main_table = pd.read_excel(file_path, sheet_name=first_sheet)
            self.main_columns = list(self.main_table.columns)
            
            message = f"主表格加载成功: {os.path.basename(file_path)}\n"
            if len(sheet_names) > 1:
                message += f"检测到 {len(sheet_names)} 个工作表，使用第一个: '{first_sheet}'\n"
            message += f"行数: {len(self.main_table)}, 列数: {len(self.main_columns)}\n"
            message += f"列名: {', '.join(self.main_columns[:5])}{'...' if len(self.main_columns) > 5 else ''}"
            
            return True, message
        except Exception as e:
            return False, f"主表格加载失败: {str(e)}"
    
    def set_key_column(self, column_index):
        """设置用于行匹配的关键列索引"""
        if column_index < 0 or column_index >= len(self.main_columns):
            return False, "列索引超出范围"
        
        self.key_column_index = column_index
        key_column_name = self.main_columns[column_index]
        return True, f"已设置关键列: 第{column_index + 1}列 '{key_column_name}'"
    
    def set_excluded_columns(self, excluded_cols):
        """设置要排除的列"""
        self.excluded_columns = set(excluded_cols) if excluded_cols else set()
        return True, f"已设置排除列: {list(self.excluded_columns)}" if self.excluded_columns else "已清空排除列"
    
    def _normalize_value(self, value):
        """标准化值，处理数据类型不一致但值相同的情况"""
        if pd.isna(value):
            return None
        
        # 转换为字符串并去掉首尾空格
        str_value = str(value).strip()
        
        # 尝试转换为数字（处理数值类型不一致的情况）
        try:
            # 如果是整数形式的浮点数，转换为整数
            if '.' in str_value and str_value.replace('.', '').isdigit():
                float_val = float(str_value)
                if float_val.is_integer():
                    return int(float_val)
                return float_val
            # 如果是纯数字，转换为整数
            elif str_value.isdigit():
                return int(str_value)
            # 尝试转换为浮点数
            elif str_value.replace('.', '').replace('-', '').isdigit():
                return float(str_value)
        except:
            pass
        
        return str_value
    
    def _values_are_different(self, val1, val2):
        """判断两个值是否不同，处理各种边界情况"""
        norm_val1 = self._normalize_value(val1)
        norm_val2 = self._normalize_value(val2)
        
        # 都是None/NaN
        if norm_val1 is None and norm_val2 is None:
            return False
        
        # 一个是None/NaN，另一个不是
        if (norm_val1 is None) != (norm_val2 is None):
            return True
        
        # 都不是None/NaN，比较值
        return norm_val1 != norm_val2
    
    def compare_tables(self, compare_files, output_dir, progress_callback=None):
        """执行表格对比 - 改进版本"""
        if self.main_table is None:
            return False, "请先加载主表格"
        
        if not compare_files:
            return False, "请选择待对比文件"
        
        key_column_name = self.main_columns[self.key_column_index]
        
        if progress_callback:
            progress_callback(f"开始对比，关键列: 第{self.key_column_index + 1}列 '{key_column_name}'")
        
        all_difference_records = []
        total_files = len(compare_files)
        total_differences = 0
        
        for file_idx, compare_file_path in enumerate(compare_files, 1):
            compare_file_name = os.path.splitext(os.path.basename(compare_file_path))[0]
            
            if progress_callback:
                progress_callback(f"[{file_idx}/{total_files}] 处理文件: {os.path.basename(compare_file_path)}")
            
            try:
                # 读取对比文件
                excel_file = pd.ExcelFile(compare_file_path)
                sheet_names = excel_file.sheet_names
                
                for sheet_name in sheet_names:
                    if progress_callback:
                        progress_callback(f"  处理工作表: {sheet_name}")
                    
                    try:
                        compare_df = pd.read_excel(compare_file_path, sheet_name=sheet_name)
                        
                        # 检查关键列是否存在
                        if key_column_name not in compare_df.columns:
                            if progress_callback:
                                progress_callback(f"    ⚠ 跳过：关键列 '{key_column_name}' 不存在")
                            continue
                        
                        # 找到两个表格的共同列
                        common_columns = list(set(self.main_table.columns) & set(compare_df.columns))
                        
                        # 排除用户指定的列
                        if self.excluded_columns:
                            common_columns = [col for col in common_columns if col not in self.excluded_columns]
                            if progress_callback:
                                progress_callback(f"    🚫 已排除用户指定的列: {list(self.excluded_columns & set(self.main_table.columns) & set(compare_df.columns))}")
                        
                        if not common_columns:
                            if progress_callback:
                                progress_callback(f"    ⚠ 跳过：没有共同列")
                            continue
                        
                        # 添加详细的共同列信息日志
                        if progress_callback:
                            main_cols = set(self.main_table.columns)
                            compare_cols = set(compare_df.columns)
                            only_in_main = main_cols - compare_cols
                            only_in_compare = compare_cols - main_cols
                            
                            progress_callback(f"    📋 主表格列数: {len(main_cols)}")
                            progress_callback(f"    📋 对比表格列数: {len(compare_cols)}")
                            progress_callback(f"    ✅ 共同列数: {len(common_columns)} - {common_columns}")
                            if only_in_main:
                                progress_callback(f"    ⚠ 主表格独有列: {list(only_in_main)}")
                            if only_in_compare:
                                progress_callback(f"    ⚠ 对比表格独有列: {list(only_in_compare)}")
                        
                        # 按关键列建立索引
                        main_indexed = self.main_table.set_index(key_column_name)
                        compare_indexed = compare_df.set_index(key_column_name)
                        
                        # 找到需要对比的行（关键列值相同的行）
                        common_keys = set(main_indexed.index) & set(compare_indexed.index)
                        
                        if not common_keys:
                            if progress_callback:
                                progress_callback(f"    ⚠ 跳过：没有匹配的关键列值")
                            continue
                        
                        sheet_differences = 0
                        
                        # 逐行对比
                        for key in common_keys:
                            main_row = main_indexed.loc[key]
                            compare_row = compare_indexed.loc[key]
                            
                            # 检查共同列是否有差异
                            row_has_differences = False
                            different_columns = []
                            
                            for col in common_columns:
                                if col == key_column_name:
                                    continue  # 跳过关键列本身
                                
                                main_value = main_row[col] if col in main_row.index else None
                                compare_value = compare_row[col] if col in compare_row.index else None
                                
                                if self._values_are_different(main_value, compare_value):
                                    row_has_differences = True
                                    different_columns.append(col)
                            
                            # 如果有差异，记录到结果中
                            if row_has_differences:
                                # 创建主表格记录（上一行）
                                main_record = {}
                                main_record[key_column_name] = key
                                for col in common_columns:
                                    if col != key_column_name:
                                        main_value = main_row[col] if col in main_row.index else ''
                                        main_record[col] = main_value
                                main_record['差异列'] = ', '.join(different_columns)
                                main_record['数据来源'] = f'{self.main_file_name}_主表格'
                                all_difference_records.append(main_record)
                                
                                # 创建对比表格记录（下一行）
                                compare_record = {}
                                compare_record[key_column_name] = key
                                for col in common_columns:
                                    if col != key_column_name:
                                        compare_value = compare_row[col] if col in compare_row.index else ''
                                        compare_record[col] = compare_value
                                compare_record['差异列'] = ', '.join(different_columns)
                                compare_record['数据来源'] = f'{compare_file_name}_{sheet_name}_对比表格'
                                all_difference_records.append(compare_record)
                                
                                sheet_differences += 1
                        
                        if progress_callback:
                            progress_callback(f"    ✓ 完成，发现 {sheet_differences} 个差异")
                        
                        total_differences += sheet_differences
                        
                    except Exception as e:
                        if progress_callback:
                            progress_callback(f"    ✗ 工作表处理失败: {str(e)}")
                        continue
                
            except Exception as e:
                if progress_callback:
                    progress_callback(f"  ✗ 文件处理失败: {str(e)}")
                continue
        
        # 生成结果文件
        if not all_difference_records:
            return True, "对比完成！未发现任何差异。"
        
        return self._save_difference_results(all_difference_records, output_dir, total_differences, progress_callback)
    
    def _save_difference_results(self, difference_records, output_dir, total_differences, progress_callback=None):
        """保存差异结果到Excel文件"""
        try:
            if progress_callback:
                progress_callback("正在生成差异报告...")
            
            # 转换为DataFrame
            diff_df = pd.DataFrame(difference_records)
            
            # 严格控制输出列：只保留真正需要的列
            key_column_name = self.main_columns[self.key_column_index]
            
            # 确定最终输出的列顺序：关键列 + 实际的共同数据列 + 差异列 + 数据来源
            if not difference_records:
                return True, "对比完成！未发现任何差异。"
            
            # 从第一条记录中获取所有列名，然后过滤
            all_columns_in_records = list(difference_records[0].keys())
            
            # 移除系统添加的列
            system_columns = ['差异列', '数据来源']
            data_columns = [col for col in all_columns_in_records if col not in system_columns]
            
            # 确保关键列在第一位
            if key_column_name in data_columns:
                data_columns.remove(key_column_name)
            final_columns = [key_column_name] + sorted(data_columns) + system_columns
            
            # 重新排列DataFrame的列顺序，只保留需要的列
            diff_df = diff_df.reindex(columns=final_columns)
            
            # 按关键列排序
            if key_column_name in diff_df.columns:
                diff_df = diff_df.sort_values(by=key_column_name)
            
            # 生成文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"差异报告_{timestamp}.xlsx"
            output_path = os.path.join(output_dir, output_filename)
            
            # 保存到Excel并设置样式
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                diff_df.to_excel(writer, sheet_name='差异报告', index=False)
                
                # 获取工作表并设置样式
                worksheet = writer.sheets['差异报告']
                
                # 设置列宽
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
                
                # 设置表头样式
                header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                
                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                
                # 高亮差异数据
                red_font = Font(color="FF0000", bold=True)
                
                # 只设置主表格的行背景色（浅蓝色），对比表格使用正常白色背景
                main_table_fill = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")  # 浅蓝色
                
                # 遍历所有数据行，标记差异和应用背景色
                for row_idx in range(2, len(diff_df) + 2):
                    # 获取数据来源信息
                    source_value = None
                    diff_cols_value = None
                    
                    for col_idx, col_name in enumerate(diff_df.columns, 1):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        
                        if col_name == '数据来源':
                            source_value = str(cell.value) if cell.value else ''
                        elif col_name == '差异列':
                            diff_cols_value = str(cell.value) if cell.value else ''
                    
                    # 只为主表格数据设置浅蓝色背景，对比表格数据使用默认白色背景
                    if source_value and '主表格' in source_value:
                        # 为整行应用浅蓝色背景
                        for col_idx in range(1, len(diff_df.columns) + 1):
                            worksheet.cell(row=row_idx, column=col_idx).fill = main_table_fill
                    # 对比表格行不设置背景色，保持默认白色
                    
                    # 标记差异列的数据
                    if diff_cols_value and diff_cols_value != 'nan':
                        diff_column_names = diff_cols_value.split(', ')
                        
                        for col_idx, col_name in enumerate(diff_df.columns, 1):
                            if col_name in diff_column_names:
                                cell = worksheet.cell(row=row_idx, column=col_idx)
                                cell.font = red_font
            
            result_message = f"""差异对比完成！
结果保存至: {output_path}
发现差异记录数: {total_differences}

对比说明:
- 关键列: 第{self.key_column_index + 1}列 '{self.main_columns[self.key_column_index]}'
- 比对方式: 只比对两表共有的列，严格排除独有列
- 排列方式: 竖排排列，主表格和对比表格数据分行显示
- 浅蓝色行: 主表格数据
- 白色行: 对比表格数据
- 红色粗体: 存在差异的数据列
- 每两行为一组差异对比记录"""
            
            return True, result_message
            
        except Exception as e:
            return False, f"保存差异报告失败: {str(e)}"
    
    def merge_files(self, compare_files, output_dir, progress_callback=None):
        """合并多个Excel文件"""
        if not compare_files:
            return False, "没有文件需要合并"
            
        try:
            merged_data = []
            total_files = len(compare_files)
            
            for i, file_path in enumerate(compare_files, 1):
                if progress_callback:
                    progress_callback(f"[{i}/{total_files}] 合并文件: {os.path.basename(file_path)}")
                
                try:
                    # 读取所有工作表
                    excel_file = pd.ExcelFile(file_path)
                    sheet_names = excel_file.sheet_names
                    
                    for sheet_name in sheet_names:
                        df = pd.read_excel(file_path, sheet_name=sheet_name)
                        
                        # 添加来源标识列
                        df['数据来源_文件'] = os.path.splitext(os.path.basename(file_path))[0]
                        df['数据来源_工作表'] = sheet_name
                        
                        merged_data.append(df)
                        
                        if progress_callback:
                            progress_callback(f"  ✓ 合并工作表: {sheet_name} ({len(df)} 行)")
                            
                except Exception as e:
                    if progress_callback:
                        progress_callback(f"  ✗ 文件处理失败: {str(e)}")
                    continue
            
            if not merged_data:
                return False, "没有成功读取到任何数据"
            
            # 合并所有数据
            if progress_callback:
                progress_callback("正在合并所有数据...")
            
            merged_df = pd.concat(merged_data, ignore_index=True)
            
            # 保存合并结果
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"合并结果_{timestamp}.xlsx"
            output_path = os.path.join(output_dir, output_filename)
            
            merged_df.to_excel(output_path, index=False)
            
            if progress_callback:
                progress_callback(f"✓ 合并完成，共合并 {len(merged_df)} 行数据")
            
            return True, f"文件合并完成！\n结果保存至: {output_path}\n总行数: {len(merged_df)}\n总列数: {len(merged_df.columns)}"
            
        except Exception as e:
            return False, f"合并过程中发生错误: {str(e)}"


class ImprovedExcelCompareGUI:
    """改进版Excel对比工具GUI界面"""
    
    def __init__(self, root):
        self.root = root
        self.engine = ImprovedExcelCompareEngine()
        self.compare_files = []
        
        self.setup_ui()
        
    def setup_ui(self):
        """设置用户界面"""
        self.root.title("Excel表格对比工具 - 改进版 v3.0")
        self.root.geometry("950x850")
        
        # 设置主题
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except:
            pass
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="Excel表格对比工具 - 改进版 v3.0", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 功能说明
        desc_text = """功能特点：
• 基于指定列进行精确行匹配    • 只比对两表共有的列，忽略独有列
• 智能处理数据类型差异        • 竖排排列，差异数据红色粗体高亮显示"""
        desc_label = ttk.Label(main_frame, text=desc_text, font=("Arial", 9), 
                              foreground="darkblue")
        desc_label.grid(row=1, column=0, columnspan=3, pady=(0, 15))
        
        # 1. 主表格选择
        ttk.Label(main_frame, text="1. 选择主对比表格:", 
                 font=("Arial", 11, "bold")).grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        
        main_file_frame = ttk.Frame(main_frame)
        main_file_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        main_file_frame.columnconfigure(0, weight=1)
        
        self.main_file_var = tk.StringVar()
        ttk.Entry(main_file_frame, textvariable=self.main_file_var, 
                 state="readonly", font=("Arial", 10)).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(main_file_frame, text="选择文件", 
                  command=self.select_main_file).grid(row=0, column=1)
        
        # 2. 关键列选择
        ttk.Label(main_frame, text="2. 选择用于行匹配的关键列:", 
                 font=("Arial", 11, "bold")).grid(row=4, column=0, sticky=tk.W, pady=(0, 5))
        
        key_column_frame = ttk.Frame(main_frame)
        key_column_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        key_column_frame.columnconfigure(1, weight=1)
        
        ttk.Label(key_column_frame, text="关键列:").grid(row=0, column=0, padx=(0, 10))
        self.key_column_var = tk.StringVar()
        self.key_column_combo = ttk.Combobox(key_column_frame, textvariable=self.key_column_var,
                                           state="readonly", font=("Arial", 10))
        self.key_column_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        self.key_column_combo.bind('<<ComboboxSelected>>', self.on_key_column_selected)
        
        # 显示关键列信息
        self.key_column_info = tk.StringVar(value="请先选择主表格")
        ttk.Label(key_column_frame, textvariable=self.key_column_info, 
                 foreground="darkgreen", font=("Arial", 9)).grid(row=1, column=0, columnspan=2, 
                                                                  sticky=tk.W, pady=(5, 0))
        
        # 2.5. 排除列选择（高级功能）
        exclude_frame = ttk.LabelFrame(main_frame, text="高级选项：排除特定列", padding="10")
        exclude_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 15))
        exclude_frame.columnconfigure(0, weight=1)
        
        ttk.Label(exclude_frame, text="如果结果中出现了不想要的列，可以在此排除：", 
                 font=("Arial", 9), foreground="darkblue").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # 添加提示文字
        ttk.Label(exclude_frame, text="输入格式：参考成本,安全库存,备注（用英文逗号分隔）", 
                 font=("Arial", 8), foreground="gray").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        
        exclude_input_frame = ttk.Frame(exclude_frame)
        exclude_input_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        exclude_input_frame.columnconfigure(0, weight=1)
        
        self.exclude_columns_var = tk.StringVar()
        self.exclude_entry = ttk.Entry(exclude_input_frame, textvariable=self.exclude_columns_var, 
                                       font=("Arial", 10))
        self.exclude_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(exclude_input_frame, text="设置排除", 
                  command=self.set_exclude_columns).grid(row=0, column=1)
        
        self.exclude_info = tk.StringVar(value="当前无排除列")
        ttk.Label(exclude_frame, textvariable=self.exclude_info, 
                 foreground="darkorange", font=("Arial", 9)).grid(row=3, column=0, sticky=tk.W, pady=(5, 0))
        
        # 3. 待对比文件
        ttk.Label(main_frame, text="3. 添加待对比文件:", 
                 font=("Arial", 11, "bold")).grid(row=7, column=0, sticky=tk.W, pady=(0, 5))
        
        files_frame = ttk.Frame(main_frame)
        files_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(0, weight=1)
        
        # 文件列表
        list_frame = ttk.Frame(files_frame)
        list_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # 创建Treeview来显示文件列表
        self.files_tree = ttk.Treeview(list_frame, columns=('file',), show='headings', height=6)
        self.files_tree.heading('file', text='待对比文件列表')
        self.files_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 滚动条
        files_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.files_tree.yview)
        files_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.files_tree.configure(yscrollcommand=files_scrollbar.set)
        
        # 文件操作按钮
        files_buttons_frame = ttk.Frame(files_frame)
        files_buttons_frame.grid(row=0, column=1, sticky=(tk.N))
        
        ttk.Button(files_buttons_frame, text="添加文件", 
                  command=self.add_compare_files).grid(row=0, column=0, pady=(0, 5))
        ttk.Button(files_buttons_frame, text="移除选中", 
                  command=self.remove_selected_file).grid(row=1, column=0, pady=(0, 5))
        ttk.Button(files_buttons_frame, text="清空列表", 
                  command=self.clear_files).grid(row=2, column=0)
        
        # 4. 输出目录
        ttk.Label(main_frame, text="4. 选择输出目录:", 
                 font=("Arial", 11, "bold")).grid(row=9, column=0, sticky=tk.W, pady=(0, 5))
        
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        output_frame.columnconfigure(0, weight=1)
        
        self.output_dir_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_dir_var, 
                 state="readonly", font=("Arial", 10)).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(output_frame, text="选择目录", 
                  command=self.select_output_dir).grid(row=0, column=1)
        
        # 5. 操作按钮
        ttk.Label(main_frame, text="5. 执行对比:", 
                 font=("Arial", 11, "bold")).grid(row=11, column=0, sticky=tk.W, pady=(0, 5))
        
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=12, column=0, columnspan=3, pady=(0, 15))
        
        self.compare_button = ttk.Button(action_frame, text="开始对比", 
                                        command=self.start_compare, 
                                        style="Accent.TButton")
        self.compare_button.grid(row=0, column=0, padx=(0, 10))
        
        ttk.Button(action_frame, text="合并文件", 
                  command=self.merge_files).grid(row=0, column=1)
        
        # 6. 日志区域
        ttk.Label(main_frame, text="6. 处理日志:", 
                 font=("Arial", 11, "bold")).grid(row=13, column=0, sticky=tk.W, pady=(0, 5))
        
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=14, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, font=("Consolas", 9))
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 9))
        status_bar.grid(row=15, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 设置行列权重
        main_frame.rowconfigure(8, weight=1)
        main_frame.rowconfigure(14, weight=1)
        
        # 初始化日志
        self.log("Excel表格对比工具 - 改进版 v3.1 已启动")
        self.log("功能说明：基于指定列精确匹配，只比对共有列，智能处理数据类型差异")
        self.log("高级功能：支持手动排除不想要的列")
    
    def log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_status(self, message):
        """更新状态栏"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def select_main_file(self):
        """选择主对比文件"""
        file_path = filedialog.askopenfilename(
            title="选择主对比表格",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            success, message = self.engine.load_main_table(file_path)
            if success:
                self.main_file_var.set(file_path)
                self.log(f"✓ {message}")
                
                # 更新关键列选择下拉框
                columns_with_index = [f"第{i+1}列: {col}" for i, col in enumerate(self.engine.main_columns)]
                self.key_column_combo['values'] = columns_with_index
                
                # 默认选择第一列
                if columns_with_index:
                    self.key_column_combo.current(0)
                    self.on_key_column_selected(None)
                
                self.update_status("主表格加载完成")
            else:
                messagebox.showerror("错误", message)
                self.log(f"✗ {message}")
    
    def on_key_column_selected(self, event):
        """关键列选择事件"""
        if self.key_column_combo.current() >= 0:
            column_index = self.key_column_combo.current()
            success, message = self.engine.set_key_column(column_index)
            if success:
                self.key_column_info.set(message)
                self.log(f"✓ {message}")
    
    def add_compare_files(self):
        """添加待对比文件"""
        file_paths = filedialog.askopenfilenames(
            title="选择待对比文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        for file_path in file_paths:
            if file_path not in self.compare_files:
                self.compare_files.append(file_path)
                # 添加到树形视图
                self.files_tree.insert('', tk.END, values=(os.path.basename(file_path),))
                self.log(f"✓ 添加文件: {os.path.basename(file_path)}")
        
        self.update_status(f"已添加 {len(self.compare_files)} 个待对比文件")
    
    def remove_selected_file(self):
        """移除选中的文件"""
        selected_items = self.files_tree.selection()
        if not selected_items:
            messagebox.showwarning("提示", "请先选择要移除的文件")
            return
        
        for item in selected_items:
            # 获取文件名
            values = self.files_tree.item(item, 'values')
            if values:
                filename = values[0]
                # 从列表中移除
                self.compare_files = [f for f in self.compare_files if os.path.basename(f) != filename]
                # 从树形视图中移除
                self.files_tree.delete(item)
                self.log(f"✓ 移除文件: {filename}")
        
        self.update_status(f"当前有 {len(self.compare_files)} 个待对比文件")
    
    def clear_files(self):
        """清空文件列表"""
        if self.compare_files:
            result = messagebox.askyesno("确认", "确定要清空所有待对比文件吗？")
            if result:
                self.compare_files.clear()
                self.files_tree.delete(*self.files_tree.get_children())
                self.log("✓ 已清空所有待对比文件")
                self.update_status("文件列表已清空")
    
    def select_output_dir(self):
        """选择输出目录"""
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.output_dir_var.set(directory)
            self.log(f"✓ 输出目录: {directory}")
            self.update_status("输出目录已选择")
    
    def merge_files(self):
        """合并文件功能"""
        if not self.compare_files:
            messagebox.showwarning("提示", "请先添加待合并的文件")
            return
        
        if not self.output_dir_var.get():
            messagebox.showwarning("提示", "请先选择输出目录")
            return
        
        self.log("开始合并文件...")
        self.update_status("正在合并文件...")
        
        def progress_callback(message):
            self.log(message)
        
        def merge_thread():
            try:
                success, message = self.engine.merge_files(
                    self.compare_files, 
                    self.output_dir_var.get(), 
                    progress_callback
                )
                self.root.after(0, lambda: self.merge_finished(success, message))
            except Exception as e:
                self.root.after(0, lambda: self.merge_finished(False, f"合并过程异常: {str(e)}"))
        
        threading.Thread(target=merge_thread, daemon=True).start()
    
    def merge_finished(self, success, message):
        """合并完成回调"""
        if success:
            self.log(f"✓ {message}")
            self.update_status("文件合并完成")
            messagebox.showinfo("成功", message)
        else:
            self.log(f"✗ {message}")
            self.update_status("文件合并失败")
            messagebox.showerror("错误", message)
    
    def start_compare(self):
        """开始对比"""
        # 验证输入
        if self.engine.main_table is None:
            messagebox.showwarning("提示", "请先选择主对比表格")
            return
        
        if not self.compare_files:
            messagebox.showwarning("提示", "请先添加待对比文件")
            return
        
        if not self.output_dir_var.get():
            messagebox.showwarning("提示", "请先选择输出目录")
            return
        
        if self.key_column_combo.current() < 0:
            messagebox.showwarning("提示", "请先选择关键列")
            return
        
        # 禁用对比按钮
        self.compare_button.config(state="disabled")
        self.log("开始执行表格对比...")
        self.update_status("正在对比表格...")
        
        def progress_callback(message):
            self.log(message)
        
        def compare_thread():
            try:
                success, message = self.engine.compare_tables(
                    self.compare_files,
                    self.output_dir_var.get(),
                    progress_callback
                )
                self.root.after(0, lambda: self.compare_complete(success, message))
            except Exception as e:
                self.root.after(0, lambda: self.compare_complete(False, f"对比过程异常: {str(e)}"))
        
        threading.Thread(target=compare_thread, daemon=True).start()
    
    def compare_complete(self, success, message):
        """对比完成回调"""
        self.compare_button.config(state="normal")
        
        if success:
            self.log(f"✓ {message}")
            self.update_status("表格对比完成")
            messagebox.showinfo("成功", message)
        else:
            self.log(f"✗ {message}")
            self.update_status("表格对比失败")
            messagebox.showerror("错误", message)
    
    def set_exclude_columns(self):
        """设置排除列"""
        exclude_text = self.exclude_columns_var.get().strip()
        if exclude_text:
            # 解析用户输入的列名
            excluded_cols = [col.strip() for col in exclude_text.split(',') if col.strip()]
            success, message = self.engine.set_excluded_columns(excluded_cols)
            if success:
                self.exclude_info.set(f"已排除列: {', '.join(excluded_cols)}")
                self.log(f"✓ {message}")
        else:
            # 清空排除列
            success, message = self.engine.set_excluded_columns([])
            self.exclude_info.set("当前无排除列")
            self.log(f"✓ {message}")
        
        self.update_status("排除列设置已更新")


def main():
    """主函数"""
    root = tk.Tk()
    app = ImprovedExcelCompareGUI(root)
    
    # 设置窗口图标和其他属性
    try:
        root.iconname("Excel Compare")
    except:
        pass
    
    # 居中显示窗口
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    root.mainloop()


if __name__ == "__main__":
    main() 