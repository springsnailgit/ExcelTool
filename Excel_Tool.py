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
import subprocess
import platform
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
import threading
from pathlib import Path

# 设置中文确认对话框
def show_confirm_dialog(title, message):
    """显示中文确认对话框"""
    dialog = tk.Toplevel()
    dialog.title(title)
    dialog.geometry("400x150")
    dialog.resizable(True, True)  # 允许调整大小
    dialog.minsize(300, 120)  # 设置最小尺寸
    dialog.transient()
    dialog.grab_set()
    
    # 配置对话框的自适应
    dialog.columnconfigure(0, weight=1)
    dialog.rowconfigure(0, weight=1)
    dialog.rowconfigure(1, weight=0)
    
    # 居中显示
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
    y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
    dialog.geometry(f"400x150+{x}+{y}")
    
    result = tk.BooleanVar()
    result.set(False)
    
    # 消息文本
    msg_label = ttk.Label(dialog, text=message, font=("Arial", 10), wraplength=350)
    msg_label.grid(row=0, column=0, columnspan=2, padx=20, pady=20, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    # 按钮框架
    button_frame = ttk.Frame(dialog)
    button_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
    
    def on_yes():
        result.set(True)
        dialog.destroy()
    
    def on_no():
        result.set(False)
        dialog.destroy()
    
    ttk.Button(button_frame, text="是", command=on_yes).grid(row=0, column=0, padx=10)
    ttk.Button(button_frame, text="否", command=on_no).grid(row=0, column=1, padx=10)
    
    dialog.wait_window()
    return result.get()

def open_and_highlight_file(file_path):
    """跨平台打开文件所在目录并高亮显示文件"""
    try:
        # 确保文件路径是绝对路径
        abs_file_path = os.path.abspath(file_path)
        
        # 根据操作系统使用不同的命令
        system = platform.system()
        if system == "Windows":
            # Windows: 使用explorer /select命令高亮显示文件
            subprocess.run(["explorer", "/select,", abs_file_path])
        elif system == "Darwin":  # macOS
            # macOS: 使用open -R命令在Finder中高亮显示文件
            subprocess.run(["open", "-R", abs_file_path])
        elif system == "Linux":
            # Linux: 尝试多种文件管理器的高亮显示命令
            # 首先尝试使用dbus-send (适用于大多数现代Linux发行版)
            try:
                subprocess.run([
                    "dbus-send", 
                    "--session", 
                    "--dest=org.freedesktop.FileManager1", 
                    "--type=method_call", 
                    "/org/freedesktop/FileManager1", 
                    "org.freedesktop.FileManager1.ShowItems", 
                    f"array:string:file://{abs_file_path}", 
                    "string:"
                ], check=True)
            except (subprocess.CalledProcessError, FileNotFoundError):
                # 如果dbus方法失败，尝试其他文件管理器
                file_managers = ["nautilus", "dolphin", "thunar", "pcmanfm", "nemo"]
                directory = os.path.dirname(abs_file_path)
                for fm in file_managers:
                    try:
                        if fm == "nautilus":
                            subprocess.run([fm, "--select", abs_file_path], check=True)
                        elif fm == "dolphin":
                            subprocess.run([fm, "--select", abs_file_path], check=True)
                        else:
                            # 对于不支持select的文件管理器，只打开目录
                            subprocess.run([fm, directory], check=True)
                        break
                    except (subprocess.CalledProcessError, FileNotFoundError):
                        continue
                else:
                    # 如果所有文件管理器都失败，使用xdg-open打开目录
                    subprocess.run(["xdg-open", directory])
        else:
            # 备用方法，使用webbrowser打开目录
            import webbrowser
            directory = os.path.dirname(abs_file_path)
            webbrowser.open(f"file://{directory}")
        
        return True
    except Exception as e:
        print(f"打开并高亮显示文件失败: {str(e)}")
        return False

class ImprovedExcelCompareEngine:
    """改进版Excel对比引擎"""
    
    def __init__(self):
        self.main_table = None
        self.main_columns = []
        self.key_column_index = 0  # 用于行匹配的关键列索引
        self.main_file_name = ""
        self.main_sheet_name = ""  # 保存主表格工作表名称
        self.selected_columns = set()  # 用户选择的参与对比的列
        
    def reset(self):
        """重置引擎状态到初始状态"""
        self.main_table = None
        self.main_columns = []
        self.key_column_index = 0
        self.main_file_name = ""
        self.main_sheet_name = ""  # 重置主表格工作表名称
        self.selected_columns = set()
        
    def load_main_table(self, file_path):
        """加载主对比表格"""
        try:
            self.main_file_name = os.path.splitext(os.path.basename(file_path))[0]
            
            # 读取Excel文件（支持多sheet，这里取第一个）
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            
            # 使用第一个sheet作为主表格
            first_sheet = sheet_names[0]
            self.main_sheet_name = first_sheet  # 保存工作表名称
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
    
    def set_selected_columns(self, selected_cols):
        """设置要参与对比的列"""
        self.selected_columns = set(selected_cols) if selected_cols else set()
        return True, f"已设置参与对比的列: {list(self.selected_columns)}" if self.selected_columns else "已清空选中列"
    
    def get_common_columns_from_files(self, compare_files):
        """获取主表格和所有对比文件的共同列"""
        if self.main_table is None:
            return []
        
        # 从主表格开始
        common_columns = set(self.main_columns)
        
        try:
            for file_idx, file_path in enumerate(compare_files, 1):
                try:
                    # 只使用第一个工作表进行比较，与实际对比逻辑保持一致
                    excel_file = pd.ExcelFile(file_path)
                    sheet_names = excel_file.sheet_names
                    
                    if sheet_names:
                        first_sheet = sheet_names[0]
                        compare_df = pd.read_excel(file_path, sheet_name=first_sheet)
                        compare_columns = set(compare_df.columns)
                        
                        # 与当前文件的列取交集
                        common_columns = common_columns & compare_columns
                        
                        # 如果没有共同列了，提前退出
                        if not common_columns:
                            break
                        
                except Exception as e:
                    # 如果读取文件失败，跳过这个文件
                    continue
                    
            return sorted(list(common_columns))
        except Exception as e:
            return []
    
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
                        
                        # 使用用户选择的列
                        if self.selected_columns:
                            # 只使用用户选择的列
                            common_columns = [col for col in common_columns if col in self.selected_columns]
                            if progress_callback:
                                progress_callback(f"    ✅ 使用用户选择的列: {list(self.selected_columns & set(self.main_table.columns) & set(compare_df.columns))}")
                        
                        if not common_columns:
                            if progress_callback:
                                progress_callback(f"    ⚠ 跳过：没有选中的共同列")
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
                                main_record['数据来源'] = f'{self.main_file_name}_{self.main_sheet_name}_主表格'
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
            return True, "对比完成！未发现任何差异。", None
        
        return self._save_difference_results(all_difference_records, output_dir, total_differences, progress_callback)
    
    def _save_difference_results(self, difference_records, output_dir, total_differences, progress_callback=None):
        """保存差异结果到Excel文件"""
        try:
            if progress_callback:
                progress_callback("正在生成差异报告...")
            
            # 重新排序差异记录，确保蓝白蓝白的排列顺序
            if progress_callback:
                progress_callback("正在重新排列差异记录...")
            
            # 按差异组重新排列：每组包含一个主表格记录和一个对比表格记录
            sorted_records = []
            key_column_name = self.main_columns[self.key_column_index]
            
            # 按关键列值分组
            main_records = {}
            compare_records = {}
            
            for record in difference_records:
                key_value = record[key_column_name]
                source = record['数据来源']
                
                if '主表格' in source:
                    if key_value not in main_records:
                        main_records[key_value] = []
                    main_records[key_value].append(record)
                else:
                    if key_value not in compare_records:
                        compare_records[key_value] = []
                    compare_records[key_value].append(record)
            
            # 按关键列值排序，然后每组内先主表格后对比表格
            all_keys = set(main_records.keys()) | set(compare_records.keys())
            for key_value in sorted(all_keys):
                # 先添加该关键值的所有主表格记录
                if key_value in main_records:
                    sorted_records.extend(main_records[key_value])
                # 再添加该关键值的所有对比表格记录
                if key_value in compare_records:
                    sorted_records.extend(compare_records[key_value])
            
            # 转换为DataFrame
            diff_df = pd.DataFrame(sorted_records)
            
            # 严格控制输出列：只保留真正需要的列
            
            # 确定最终输出的列顺序：关键列 + 实际选中的数据列 + 差异列 + 数据来源
            if not sorted_records:
                return True, "对比完成！未发现任何差异。", None
            
            # 从第一条记录中获取所有列名，然后过滤
            all_columns_in_records = list(sorted_records[0].keys())
            
            # 移除系统添加的列
            system_columns = ['差异列', '数据来源']
            data_columns = [col for col in all_columns_in_records if col not in system_columns]
            
            # 确保关键列在第一位
            if key_column_name in data_columns:
                data_columns.remove(key_column_name)
            final_columns = [key_column_name] + sorted(data_columns) + system_columns
            
            # 重新排列DataFrame的列顺序，只保留需要的列
            diff_df = diff_df.reindex(columns=final_columns)
            
            # 注意：不再按关键列重新排序，保持我们已经排序好的蓝白蓝白顺序
            
            # 生成文件名：差异报告+日期+序号
            date_str = datetime.now().strftime("%Y%m%d")
            
            # 查找同日期已存在的报告数量，生成序号
            base_filename = f"差异报告_{date_str}"
            counter = 1
            output_filename = f"{base_filename}_{counter:03d}.xlsx"
            output_path = os.path.join(output_dir, output_filename)
            
            # 如果文件已存在，递增序号
            while os.path.exists(output_path):
                counter += 1
                output_filename = f"{base_filename}_{counter:03d}.xlsx"
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
                
                # 设置主表格的行背景色（浅蓝色），对比表格使用正常白色背景
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
                    
                    # 为主表格数据设置浅蓝色背景，对比表格数据使用默认白色背景
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
- 比对方式: 只对比用户选择的列，关键列为必选项
- 排列方式: 竖排排列，主表格和对比表格数据分行显示，蓝白蓝白排列
- 浅蓝色行: 主表格数据
- 白色行: 对比表格数据
- 红色粗体: 存在差异的数据列
- 每两行为一组差异对比记录
- 数据来源: 包含文件名和工作表名称
- 文件命名: 差异报告_日期_序号格式"""
            
            return True, result_message, output_path
            
        except Exception as e:
            return False, f"保存差异报告失败: {str(e)}", None


class ImprovedExcelCompareGUI:
    """改进版Excel对比工具GUI界面"""
    
    def __init__(self, root):
        self.root = root
        self.engine = ImprovedExcelCompareEngine()
        self.compare_files = []
        
        self.setup_ui()
        
    def setup_ui(self):
        """设置用户界面"""
        self.root.title("Excel表格对比工具 - 改进版 v4.1.1")
        self.root.geometry("1100x850")  # 调整窗口大小以适应新布局
        
        # 设置最小窗口大小，确保界面不会过小
        self.root.minsize(900, 700)
        
        # 设置主题
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except:
            pass
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置根窗口的自适应
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # 配置主框架的自适应 - 每列都有相等权重，可以灵活调整
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="Excel表格对比工具 - 改进版 v4.1.1", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 功能说明
        desc_text = """功能特点：
• 基于指定列进行精确行匹配    • 智能选择参与对比的列，关键列必选
• 智能处理数据类型差异        • 竖排排列，差异数据红色粗体高亮显示
• 弹出窗口选择比对列         • 自动生成序号的差异报告
• 并排布局优化界面           • 精简操作流程"""
        desc_label = ttk.Label(main_frame, text=desc_text, font=("Arial", 9), 
                              foreground="darkblue")
        desc_label.grid(row=1, column=0, columnspan=2, pady=(0, 15))
        
        # 第一行：主表格选择 | 关键列选择（并排）
        # 左侧：主表格选择
        main_file_frame = ttk.LabelFrame(main_frame, text="1. 选择主对比表格", padding="10")
        main_file_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10), pady=(0, 15))
        main_file_frame.columnconfigure(0, weight=1)
        
        self.main_file_var = tk.StringVar()
        ttk.Entry(main_file_frame, textvariable=self.main_file_var, 
                 state="readonly", font=("Arial", 10)).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10), pady=(0, 10))
        ttk.Button(main_file_frame, text="选择文件", 
                  command=self.select_main_file).grid(row=0, column=1, pady=(0, 10))
        
        # 右侧：关键列选择
        key_column_frame = ttk.LabelFrame(main_frame, text="2. 选择用于行匹配的关键列", padding="10")
        key_column_frame.grid(row=2, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        key_column_frame.columnconfigure(1, weight=1)
        
        ttk.Label(key_column_frame, text="关键列:").grid(row=0, column=0, padx=(0, 10))
        self.key_column_var = tk.StringVar()
        self.key_column_combo = ttk.Combobox(key_column_frame, textvariable=self.key_column_var,
                                           state="readonly", font=("Arial", 10))
        self.key_column_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 0), pady=(0, 5))
        self.key_column_combo.bind('<<ComboboxSelected>>', self.on_key_column_selected)
        
        # 显示关键列信息
        self.key_column_info = tk.StringVar(value="请先选择主表格")
        ttk.Label(key_column_frame, textvariable=self.key_column_info, 
                 foreground="darkgreen", font=("Arial", 9), wraplength=200).grid(row=1, column=0, columnspan=2, 
                                                                  sticky=tk.W, pady=(5, 0))
        
        # 第二行：待对比文件 | 比对设置（并排）
        # 左侧：待对比文件
        files_frame = ttk.LabelFrame(main_frame, text="3. 添加待对比文件", padding="10")
        files_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10), pady=(0, 15))
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
        
        # 右侧：比对设置（竖排布局）
        settings_frame = ttk.LabelFrame(main_frame, text="4. 比对设置", padding="10")
        settings_frame.grid(row=3, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        settings_frame.columnconfigure(0, weight=1)
        # 优化权重分配：两个子模块都参与伸缩，但输出目录有最小高度保护
        settings_frame.rowconfigure(0, weight=2)  # 比对列选择区域，权重较大
        settings_frame.rowconfigure(1, weight=1, minsize=80)  # 输出目录区域，权重较小但有最小高度保护
        
        # 上部：比对列选择
        exclude_frame = ttk.LabelFrame(settings_frame, text="比对列选择", padding="8")
        exclude_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        exclude_frame.columnconfigure(0, weight=1)
        # 优化内部元素布局：让状态信息能够更好地适应空间变化
        exclude_frame.rowconfigure(0, weight=0)  # 说明文字，固定高度
        exclude_frame.rowconfigure(1, weight=0)  # 按钮，固定高度  
        exclude_frame.rowconfigure(2, weight=1)  # 状态信息，充分利用可用空间
        
        ttk.Label(exclude_frame, text="选择哪些列参与对比（关键列为必选项）：", 
                 font=("Arial", 9), foreground="darkblue").grid(row=0, column=0, sticky=tk.W, pady=(0, 8))
        
        # 只保留选择比对列按钮
        self.exclude_expand_button = ttk.Button(exclude_frame, text="选择比对列", 
                                               command=self.open_column_selection_window)
        self.exclude_expand_button.grid(row=1, column=0, pady=(0, 8))
        
        # 状态信息
        self.exclude_info = tk.StringVar(value="当前无选中列 - 请先选择主表格和待对比文件")
        ttk.Label(exclude_frame, textvariable=self.exclude_info, 
                 foreground="darkorange", font=("Arial", 9), wraplength=200).grid(row=2, column=0, sticky=(tk.W, tk.N), pady=(0, 8))
        
        # 下部：输出目录选择
        output_frame = ttk.LabelFrame(settings_frame, text="输出目录", padding="8")
        output_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        output_frame.columnconfigure(0, weight=1)
        # 输出目录区域内部布局：确保控件在任何大小下都可见可操作
        output_frame.rowconfigure(0, weight=0)  # 说明文字，固定高度
        output_frame.rowconfigure(1, weight=1)  # 输入控件区域，可适应调整但保持可用性
        
        ttk.Label(output_frame, text="选择差异报告的保存位置：", 
                 font=("Arial", 9), foreground="darkblue").grid(row=0, column=0, sticky=tk.W, pady=(0, 8))
        
        output_controls_frame = ttk.Frame(output_frame)
        output_controls_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        output_controls_frame.columnconfigure(0, weight=1)
        
        self.output_dir_var = tk.StringVar()
        ttk.Entry(output_controls_frame, textvariable=self.output_dir_var, 
                 state="readonly", font=("Arial", 10)).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(output_controls_frame, text="选择目录", 
                  command=self.select_output_dir).grid(row=0, column=1)
        
        # 保留折叠功能（隐藏）但不显示在主界面中
        self.exclude_checkboxes_frame = ttk.Frame(main_frame)
        self.exclude_canvas = tk.Canvas(self.exclude_checkboxes_frame, height=0)
        self.exclude_scrollbar = ttk.Scrollbar(self.exclude_checkboxes_frame, orient="vertical", 
                                              command=self.exclude_canvas.yview)
        self.exclude_scrollable_frame = ttk.Frame(self.exclude_canvas)
        
        self.exclude_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.exclude_canvas.configure(scrollregion=self.exclude_canvas.bbox("all"))
        )
        
        self.exclude_canvas.create_window((0, 0), window=self.exclude_scrollable_frame, anchor="nw")
        self.exclude_canvas.configure(yscrollcommand=self.exclude_scrollbar.set)
        
        # 初始状态变量
        self.exclude_columns_expanded = False
        self.exclude_column_vars = {}  # 存储每列的选择状态
        self.selected_columns = set()  # 存储选中的列
        
        # 5. 执行对比
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=4, column=0, columnspan=2, pady=(0, 15))
        
        ttk.Label(action_frame, text="5. 执行对比:", 
                 font=("Arial", 11, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        
        self.compare_button = ttk.Button(action_frame, text="开始对比", 
                                        command=self.start_compare, 
                                        style="Accent.TButton")
        self.compare_button.grid(row=0, column=1, padx=(0, 10))
        
        ttk.Button(action_frame, text="重置", 
                  command=self.reset).grid(row=0, column=2)
        
        # 6. 日志区域
        ttk.Label(main_frame, text="6. 处理日志:", 
                 font=("Arial", 11, "bold")).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, font=("Consolas", 9))
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 9))
        status_bar.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 完善自适应配置 - 设置行列权重以确保各区域能够自适应调整
        # 行权重配置：让需要伸缩的区域有合适的权重
        main_frame.rowconfigure(0, weight=0)  # 标题行，固定高度
        main_frame.rowconfigure(1, weight=0)  # 功能说明行，固定高度
        main_frame.rowconfigure(2, weight=0)  # 主表格和关键列选择行，最小高度
        main_frame.rowconfigure(3, weight=3)  # 文件列表和设置区域，主要伸缩区域，给予最大权重
        main_frame.rowconfigure(4, weight=0)  # 执行对比按钮行，固定高度
        main_frame.rowconfigure(5, weight=0)  # 日志标题行，固定高度
        main_frame.rowconfigure(6, weight=2)  # 日志区域，次要伸缩区域
        main_frame.rowconfigure(7, weight=0)  # 状态栏，固定高度
        
        # 初始化日志
        self.log("Excel表格对比工具 - 改进版 v4.1.1 已启动")
        self.log("新功能：比对列选择，弹出窗口智能选择，关键列必选机制")
        self.log("新特性：差异报告自动序号命名，只对比选中的列")
        self.log("界面优化：主表格和关键列并排，待对比文件和比对设置并排")
    
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
                
                # 更新排除列选择的可用性
                self.exclude_info.set(f"可选择列数: {len(self.engine.main_columns)} - 点击'选择比对列'开始设置")
                
                # 如果当前展开着排除列选择，更新复选框列表
                if self.exclude_columns_expanded:
                    self._update_exclude_checkboxes()
                
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
            result = show_confirm_dialog("确认清空", "确定要清空所有待对比文件吗？")
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
                success, message, output_path = self.engine.compare_tables(
                    self.compare_files,
                    self.output_dir_var.get(),
                    progress_callback
                )
                self.root.after(0, lambda: self.compare_complete(success, message, output_path))
            except Exception as e:
                self.root.after(0, lambda: self.compare_complete(False, f"对比过程异常: {str(e)}", None))
        
        threading.Thread(target=compare_thread, daemon=True).start()
    
    def compare_complete(self, success, message, output_path):
        """对比完成回调"""
        self.compare_button.config(state="normal")
        
        if success:
            self.log(f"✓ {message}")
            self.update_status("表格对比完成")
            messagebox.showinfo("成功", message)
            
            # 如果生成了差异报告文件，自动打开保存目录
            if output_path and os.path.exists(output_path):
                self.log("正在打开差异报告保存目录并高亮显示文件...")
                if open_and_highlight_file(output_path):
                    self.log("✓ 已自动打开差异报告保存目录并高亮显示差异报告文件")
                else:
                    self.log("✗ 打开目录或高亮显示文件失败，请手动查看文件")
            
            return output_path
        else:
            self.log(f"✗ {message}")
            self.update_status("表格对比失败")
            messagebox.showerror("错误", message)
            return None
    
    def open_column_selection_window(self):
        """打开列选择弹出窗口"""
        if self.engine.main_table is None:
            messagebox.showwarning("提示", "请先选择主表格")
            return
            
        if not self.compare_files:
            messagebox.showwarning("提示", "请先添加待对比文件")
            return
        
        # 获取共同列
        self.log(f"开始检测共同列...")
        self.log(f"主表格列数: {len(self.engine.main_columns)}")
        self.log(f"主表格列名: {self.engine.main_columns[:5]}{'...' if len(self.engine.main_columns) > 5 else ''}")
        self.log(f"待对比文件数: {len(self.compare_files)}")
        
        common_columns = self.engine.get_common_columns_from_files(self.compare_files)
        
        self.log(f"检测到共同列数: {len(common_columns)}")
        if common_columns:
            self.log(f"共同列名: {common_columns[:5]}{'...' if len(common_columns) > 5 else ''}")
        
        if not common_columns:
            messagebox.showwarning("提示", 
                                 f"主表格与对比文件没有共同列\n\n" +
                                 f"主表格列数: {len(self.engine.main_columns)}\n" +
                                 f"对比文件数: {len(self.compare_files)}\n\n" +
                                 f"请检查：\n" +
                                 f"1. 文件格式是否正确\n" +
                                 f"2. 列名是否完全一致\n" +
                                 f"3. 是否存在空格或特殊字符差异")
            return
        
        # 创建弹出窗口
        self.column_window = tk.Toplevel(self.root)
        self.column_window.title("选择参与对比的列")
        self.column_window.geometry("500x400")
        self.column_window.resizable(True, True)
        
        # 设置弹出窗口的最小大小
        self.column_window.minsize(400, 300)
        
        # 设置窗口居中
        self.column_window.transient(self.root)
        self.column_window.grab_set()
        
        # 主框架
        main_frame = ttk.Frame(self.column_window, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置弹出窗口的自适应
        self.column_window.columnconfigure(0, weight=1)
        self.column_window.rowconfigure(0, weight=1)
        
        # 配置主框架的自适应
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=0)  # 标题行，固定高度
        main_frame.rowconfigure(1, weight=0)  # 说明行，固定高度
        main_frame.rowconfigure(2, weight=1)  # 列选择区域，主要伸缩区域
        main_frame.rowconfigure(3, weight=0)  # 按钮行，固定高度
        
        # 标题
        title_label = ttk.Label(main_frame, text="选择参与对比的列", 
                               font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 15))
        
        # 说明文字
        key_column_name = self.engine.main_columns[self.engine.key_column_index]
        desc_text = f"以下是主表格与对比文件的共同列：\n关键列 '{key_column_name}' 为必选项，其他列可自由选择"
        desc_label = ttk.Label(main_frame, text=desc_text, 
                              font=("Arial", 10), foreground="darkblue")
        desc_label.grid(row=1, column=0, pady=(0, 10))
        
        # 列选择区域
        list_frame = ttk.Frame(main_frame)
        list_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # 创建滚动区域
        canvas = tk.Canvas(list_frame)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 存储复选框变量
        self.column_selection_vars = {}
        
        # 为每个共同列创建复选框
        for i, col in enumerate(common_columns):
            var = tk.BooleanVar()
            
            # 关键列默认选中且不可更改
            if col == key_column_name:
                var.set(True)
                checkbox = ttk.Checkbutton(
                    scrollable_frame,
                    text=f"第{self.engine.main_columns.index(col) + 1}列: {col} (必选)",
                    variable=var,
                    state="disabled"
                )
            else:
                # 如果之前有选择，保持选择状态
                if col in self.engine.selected_columns:
                    var.set(True)
                checkbox = ttk.Checkbutton(
                    scrollable_frame,
                    text=f"第{self.engine.main_columns.index(col) + 1}列: {col}",
                    variable=var
                )
            
            checkbox.grid(row=i, column=0, sticky=tk.W, padx=10, pady=2)
            self.column_selection_vars[col] = var
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, pady=(10, 0))
        
        ttk.Button(button_frame, text="全选", 
                  command=self.select_all_columns_in_window).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(button_frame, text="全清", 
                  command=self.clear_all_columns_in_window).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(button_frame, text="确定", 
                  command=self.confirm_column_selection).grid(row=0, column=2, padx=(0, 10))
        ttk.Button(button_frame, text="取消", 
                  command=self.column_window.destroy).grid(row=0, column=3)
    
    def select_all_columns_in_window(self):
        """在弹出窗口中全选所有列"""
        key_column_name = self.engine.main_columns[self.engine.key_column_index]
        for col, var in self.column_selection_vars.items():
            if col != key_column_name:  # 关键列已经是必选，跳过
                var.set(True)
    
    def clear_all_columns_in_window(self):
        """在弹出窗口中清空所有列（除了关键列）"""
        key_column_name = self.engine.main_columns[self.engine.key_column_index]
        for col, var in self.column_selection_vars.items():
            if col != key_column_name:  # 关键列不能取消
                var.set(False)
    
    def confirm_column_selection(self):
        """确认列选择"""
        selected_columns = [col for col, var in self.column_selection_vars.items() if var.get()]
        
        # 确保关键列被选中
        key_column_name = self.engine.main_columns[self.engine.key_column_index]
        if key_column_name not in selected_columns:
            selected_columns.append(key_column_name)
        
        # 更新引擎
        success, message = self.engine.set_selected_columns(selected_columns)
        
        if success:
            # 更新界面显示
            if selected_columns:
                self.exclude_info.set(f"已选择 {len(selected_columns)} 列参与对比")
                self.log(f"✓ {message}")
            else:
                self.exclude_info.set("当前无选中列")
                self.log(f"✓ {message}")
        
        # 关闭窗口
        self.column_window.destroy()
        self.update_status("比对列选择已更新")
    
    def reset(self):
        """重置所有输入和状态"""
        result = show_confirm_dialog("确认重置", "确定要重置所有输入和状态吗？\n这将清空所有已选择的文件和设置。")
        if not result:
            return
            
        # 重置引擎状态
        self.engine.reset()
        
        # 重置文件列表
        self.compare_files.clear()
        self.files_tree.delete(*self.files_tree.get_children())
        
        # 重置主文件选择
        self.main_file_var.set("")
        
        # 重置关键列选择
        self.key_column_combo.delete(0, tk.END)
        self.key_column_combo.set("")
        self.key_column_var.set("")
        self.key_column_info.set("请先选择主表格")
        
        # 重置排除列设置
        self.exclude_column_vars.clear()
        self.selected_columns.clear()
        # 清空复选框区域
        for widget in self.exclude_scrollable_frame.winfo_children():
            widget.destroy()
        # 隐藏复选框区域
        if self.exclude_columns_expanded:
            self.exclude_canvas.grid_remove()
            self.exclude_scrollbar.grid_remove()
            self.exclude_expand_button.config(text="选择比对列")
            self.exclude_columns_expanded = False
        self.exclude_info.set("当前无选中列 - 请先选择主表格和待对比文件")
        
        # 重置输出目录
        self.output_dir_var.set("")
        
        # 清空日志
        self.log_text.delete(1.0, tk.END)
        
        # 重新初始化日志
        self.log("Excel表格对比工具 - 改进版 v4.1.1 已重置")
        self.log("新功能：比对列选择，弹出窗口智能选择，关键列必选机制")
        self.log("新特性：差异报告自动序号命名，只对比选中的列")
        self.log("界面优化：主表格和关键列并排，待对比文件和比对设置并排")
        self.log("✓ 所有输入和状态已重置到初始状态")
        
        # 更新状态栏
        self.update_status("就绪 - 已重置所有状态")


def main():
    """主函数"""
    root = tk.Tk()
    app = ImprovedExcelCompareGUI(root)
    
    # 设置窗口图标和其他属性
    try:
        root.iconname("Excel Compare")
    except:
        pass
    
    # 添加窗口大小变化的回调，用于优化自适应显示
    def on_window_resize(event):
        """窗口大小变化时的回调函数"""
        if event.widget == root:
            # 当窗口大小变化时，确保界面元素正确更新
            root.update_idletasks()
    
    # 绑定窗口大小变化事件
    root.bind('<Configure>', on_window_resize)
    
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