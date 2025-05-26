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
import json

# 添加OpenAI导入用于DeepSeek API
try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False
    print("警告: OpenAI库未安装，DeepSeek功能将受限。请运行: pip install openai")

# 环境变量配置管理
def load_env_config():
    """加载环境变量配置"""
    config = {
        'api_key': '',
        'base_url': 'https://api.deepseek.com',
        'default_model': 'deepseek-chat',
        'timeout': 30
    }
    
    # 尝试从环境变量读取
    config['api_key'] = os.getenv('DEEPSEEK_API_KEY', '')
    config['base_url'] = os.getenv('DEEPSEEK_BASE_URL', 'https://api.deepseek.com')
    config['default_model'] = os.getenv('DEEPSEEK_DEFAULT_MODEL', 'deepseek-chat')
    
    try:
        config['timeout'] = int(os.getenv('DEEPSEEK_TIMEOUT', '30'))
    except ValueError:
        config['timeout'] = 30
    
    # 尝试从.env文件读取
    env_file_path = os.path.join(os.path.dirname(__file__), '.env')
    if os.path.exists(env_file_path):
        try:
            with open(env_file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#') and '=' in line:
                        key, value = line.split('=', 1)
                        key = key.strip()
                        value = value.strip()
                        
                        if key == 'DEEPSEEK_API_KEY':
                            config['api_key'] = value
                        elif key == 'DEEPSEEK_BASE_URL':
                            config['base_url'] = value
                        elif key == 'DEEPSEEK_DEFAULT_MODEL':
                            config['default_model'] = value
                        elif key == 'DEEPSEEK_TIMEOUT':
                            try:
                                config['timeout'] = int(value)
                            except ValueError:
                                pass
        except Exception as e:
            print(f"警告: 读取.env文件失败: {e}")
    
    return config

def create_env_file():
    """创建.env文件"""
    env_content = """# DeepSeek API 配置文件
# 请将此文件添加到 .gitignore 中以确保API Key不会被提交到版本控制系统

# DeepSeek API Key (请替换为您的真实API Key)
DEEPSEEK_API_KEY=your_api_key_here

# DeepSeek API 基础URL（可选，默认为官方地址）
DEEPSEEK_BASE_URL=https://api.deepseek.com

# 默认模型设置（可选）
DEEPSEEK_DEFAULT_MODEL=deepseek-chat

# API调用超时设置（秒）
DEEPSEEK_TIMEOUT=30

# 是否启用调试模式
DEEPSEEK_DEBUG=false
"""
    
    try:
        env_file_path = os.path.join(os.path.dirname(__file__), '.env')
        with open(env_file_path, 'w', encoding='utf-8') as f:
            f.write(env_content)
        print("✓ .env文件已创建")
        return True
    except Exception as e:
        print(f"✗ 创建.env文件失败: {e}")
        return False

# 加载环境配置
ENV_CONFIG = load_env_config()

# DeepSeek模型配置
DEEPSEEK_MODELS = {
    "deepseek-chat": {
        "name": "DeepSeek Chat",
        "description": "通用对话模型，适用于各种对话和数据分析任务",
        "endpoint": "https://api.deepseek.com"
    },
    "deepseek-reasoner": {
        "name": "DeepSeek Reasoner", 
        "description": "推理模型，专门用于复杂推理和分析任务",
        "endpoint": "https://api.deepseek.com"
    }
}

# DeepSeek API模型管理类
class DeepSeekModelManager:
    """DeepSeek模型管理器，用于获取和管理可用模型"""
    
    def __init__(self, api_key=None, base_url=None):
        # 优先使用传入参数，否则使用环境变量配置
        self.api_key = api_key or ENV_CONFIG['api_key']
        self.base_url = base_url or ENV_CONFIG['base_url']
        self.available_models = {}
        
        # 如果没有提供API Key，尝试自动获取
        if not self.api_key:
            self.api_key = ENV_CONFIG['api_key']
        
        # 创建.env文件（如果不存在）
        env_file_path = os.path.join(os.path.dirname(__file__), '.env')
        if not os.path.exists(env_file_path):
            create_env_file()
        
    def list_models(self):
        """
        列出可用的模型列表
        根据官方API文档: GET /models
        返回格式: {"object": "list", "data": [{"id": "model_id", "object": "model", "owned_by": "deepseek"}]}
        """
        if not OPENAI_AVAILABLE:
            print("警告: OpenAI库未安装，无法获取模型列表")
            return DEEPSEEK_MODELS
            
        if not self.api_key:
            print("警告: 未提供API Key，使用默认模型列表")
            return DEEPSEEK_MODELS
            
        try:
            client = OpenAI(
                api_key=self.api_key,
                base_url=self.base_url
            )
            
            # 调用API获取模型列表
            models_response = client.models.list()
            
            # 处理响应数据
            available_models = {}
            for model_data in models_response.data:
                model_id = model_data.id
                
                # 根据模型ID设置描述信息
                if model_id == "deepseek-chat":
                    description = "通用对话模型，适用于各种对话和数据分析任务"
                    name = "DeepSeek Chat"
                elif model_id == "deepseek-reasoner":
                    description = "推理模型，专门用于复杂推理和分析任务"
                    name = "DeepSeek Reasoner"
                else:
                    description = f"DeepSeek模型: {model_id}"
                    name = model_id.replace("-", " ").title()
                
                available_models[model_id] = {
                    "name": name,
                    "description": description,
                    "endpoint": self.base_url,
                    "owned_by": getattr(model_data, 'owned_by', 'deepseek')
                }
            
            self.available_models = available_models
            print(f"成功获取 {len(available_models)} 个可用模型")
            return available_models
            
        except Exception as e:
            print(f"获取模型列表失败: {str(e)}")
            print("使用默认模型列表")
            return DEEPSEEK_MODELS
    
    def get_model_info(self, model_id):
        """获取特定模型的详细信息"""
        if model_id in self.available_models:
            return self.available_models[model_id]
        elif model_id in DEEPSEEK_MODELS:
            return DEEPSEEK_MODELS[model_id]
        else:
            return {
                "name": model_id,
                "description": f"未知模型: {model_id}",
                "endpoint": self.base_url
            }
    
    def refresh_models(self):
        """刷新模型列表"""
        return self.list_models()

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

class ModelSelectionWindow:
    """DeepSeek模型选择窗口"""
    
    def __init__(self, parent, api_key=None):
        self.parent = parent
        self.selected_model = None
        self.selected_temperature = 0.7
        self.result = None
        self.api_key = api_key
        self.model_manager = DeepSeekModelManager(api_key)
        self.available_models = DEEPSEEK_MODELS  # 默认模型列表
        
    def show(self):
        """显示模型选择窗口"""
        self.window = tk.Toplevel(self.parent)
        self.window.title("选择DeepSeek模型")
        self.window.geometry("500x400")
        self.window.resizable(True, True)
        self.window.transient(self.parent)
        self.window.grab_set()
        
        # 居中显示
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"500x400+{x}+{y}")
        
        # 配置窗口
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        
        # 主框架
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="选择DeepSeek分析模型", 
                               font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # 模型选择区域
        model_frame = ttk.Frame(main_frame)
        model_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        model_frame.columnconfigure(0, weight=1)
        
        ttk.Label(model_frame, text="选择模型:", font=("Arial", 11)).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # 模型下拉框和刷新按钮
        dropdown_frame = ttk.Frame(model_frame)
        dropdown_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))
        dropdown_frame.columnconfigure(0, weight=1)
        
        self.model_var = tk.StringVar(value="deepseek-chat")
        self.model_dropdown = ttk.Combobox(
            dropdown_frame,
            textvariable=self.model_var,
            values=list(self.available_models.keys()),
            state="readonly",
            font=("Arial", 10)
        )
        self.model_dropdown.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        self.model_dropdown.bind("<<ComboboxSelected>>", self._update_description)
        
        # 刷新模型列表按钮
        refresh_button = ttk.Button(
            dropdown_frame,
            text="刷新模型",
            command=self._refresh_models,
            width=10
        )
        refresh_button.grid(row=0, column=1)
        
        # 模型描述
        desc_frame = ttk.LabelFrame(main_frame, text="模型描述", padding="10")
        desc_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        desc_frame.columnconfigure(0, weight=1)
        desc_frame.rowconfigure(1, weight=1)
        
        self.model_name_label = ttk.Label(desc_frame, text="", font=("Arial", 11, "bold"))
        self.model_name_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.model_desc_label = ttk.Label(desc_frame, text="", font=("Arial", 10), 
                                         wraplength=400, justify=tk.LEFT)
        self.model_desc_label.grid(row=1, column=0, sticky=(tk.W, tk.N), pady=(0, 10))
        
        # 参数设置
        param_frame = ttk.LabelFrame(main_frame, text="参数设置", padding="10")
        param_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        param_frame.columnconfigure(1, weight=1)
        
        ttk.Label(param_frame, text="Temperature:", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.temp_var = tk.DoubleVar(value=0.7)
        temp_scale = ttk.Scale(param_frame, from_=0.0, to=1.0, variable=self.temp_var, 
                              orient=tk.HORIZONTAL)
        temp_scale.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.temp_label = ttk.Label(param_frame, text="0.7", font=("Arial", 10))
        self.temp_label.grid(row=0, column=2)
        
        temp_scale.bind("<Motion>", self._update_temp_label)
        temp_scale.bind("<ButtonRelease-1>", self._update_temp_label)
        
        # 流式响应选项
        self.stream_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(param_frame, text="启用流式响应", 
                       variable=self.stream_var).grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))
        
        # 按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, pady=(10, 0))
        
        ttk.Button(button_frame, text="确认", command=self._confirm).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(button_frame, text="取消", command=self._cancel).grid(row=0, column=1)
        
        # 初始化描述
        self._update_description(None)
        
        # 等待窗口关闭
        self.window.wait_window()
        return self.result
    
    def _update_description(self, event):
        """更新模型描述"""
        selected_model = self.model_var.get()
        model_info = self.model_manager.get_model_info(selected_model)
        self.model_name_label.config(text=model_info["name"])
        self.model_desc_label.config(text=model_info["description"])
        
    def _refresh_models(self):
        """刷新模型列表"""
        try:
            self.model_name_label.config(text="正在刷新模型列表...")
            self.model_desc_label.config(text="请稍候...")
            self.window.update()
            
            # 获取最新的模型列表
            self.available_models = self.model_manager.list_models()
            
            # 更新下拉框选项
            self.model_dropdown['values'] = list(self.available_models.keys())
            
            # 如果当前选择的模型不在新列表中，选择第一个可用模型
            current_selection = self.model_var.get()
            if current_selection not in self.available_models:
                if self.available_models:
                    self.model_var.set(list(self.available_models.keys())[0])
                else:
                    self.model_var.set("")
            
            # 更新描述
            self._update_description(None)
            
            messagebox.showinfo("刷新完成", f"成功获取 {len(self.available_models)} 个可用模型")
            
        except Exception as e:
            error_msg = f"刷新模型列表失败: {str(e)}"
            self.model_name_label.config(text="刷新失败")
            self.model_desc_label.config(text=error_msg)
            messagebox.showerror("刷新失败", error_msg)
    
    def _update_temp_label(self, event):
        """更新温度标签"""
        temp_value = round(self.temp_var.get(), 2)
        self.temp_label.config(text=str(temp_value))
    
    def _confirm(self):
        """确认选择"""
        self.result = {
            "model": self.model_var.get(),
            "temperature": round(self.temp_var.get(), 2),
            "endpoint": DEEPSEEK_MODELS[self.model_var.get()]["endpoint"],
            "stream": self.stream_var.get()
        }
        self.window.destroy()
    
    def _cancel(self):
        """取消选择"""
        self.result = None
        self.window.destroy()

class ChatWindow:
    """DeepSeek对话分析窗口"""
    
    def __init__(self, parent, desensitized_data, api_key, model_config):
        self.parent = parent
        self.desensitized_data = desensitized_data
        self.api_key = api_key
        self.model_config = model_config
        self.chat_history = []
        
        # 准备数据上下文
        self.context = self._prepare_context()
        
    def show(self):
        """显示对话窗口"""
        self.window = tk.Toplevel(self.parent)
        self.window.title(f"DeepSeek对话分析 - {DEEPSEEK_MODELS[self.model_config['model']]['name']}")
        self.window.geometry("900x700")
        self.window.resizable(True, True)
        
        # 配置窗口
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        
        # 主框架
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # 顶部信息栏
        info_frame = ttk.Frame(main_frame)
        info_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        info_frame.columnconfigure(1, weight=1)
        
        ttk.Label(info_frame, text="当前模型:", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W)
        model_info = f"{self.model_config['model']} (Temperature: {self.model_config['temperature']})"
        ttk.Label(info_frame, text=model_info, font=("Arial", 10, "bold"), 
                 foreground="darkblue").grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        
        ttk.Button(info_frame, text="保存对话", command=self._save_chat).grid(row=0, column=2, padx=(10, 0))
        ttk.Button(info_frame, text="清空对话", command=self._clear_chat).grid(row=0, column=3, padx=(10, 0))
        
        # 对话显示区域
        chat_frame = ttk.Frame(main_frame)
        chat_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        chat_frame.columnconfigure(0, weight=1)
        chat_frame.rowconfigure(0, weight=1)
        
        self.chat_text = scrolledtext.ScrolledText(
            chat_frame, 
            wrap=tk.WORD, 
            font=("Arial", 10),
            state='disabled'
        )
        self.chat_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置文本标签样式
        self.chat_text.tag_configure("user", foreground="blue", font=("Arial", 10, "bold"))
        self.chat_text.tag_configure("assistant", foreground="darkgreen", font=("Arial", 10, "bold"))
        self.chat_text.tag_configure("system", foreground="gray", font=("Arial", 9, "italic"))
        
        # 输入区域
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))
        input_frame.columnconfigure(0, weight=1)
        
        # 输入提示
        ttk.Label(input_frame, text="输入您的问题:", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # 输入控件
        entry_frame = ttk.Frame(input_frame)
        entry_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))
        entry_frame.columnconfigure(0, weight=1)
        
        self.input_var = tk.StringVar()
        self.input_entry = ttk.Entry(entry_frame, textvariable=self.input_var, font=("Arial", 10))
        self.input_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        self.input_entry.bind("<Return>", lambda e: self._send_message())
        
        self.send_button = ttk.Button(entry_frame, text="发送", command=self._send_message)
        self.send_button.grid(row=0, column=1)
        
        # 初始化对话
        self._display_message("system", "DeepSeek助手已准备就绪！我已经了解了您的数据概况，请随时提问。")
        self._display_message("system", f"数据概况: {self.desensitized_data.shape[0]}行 x {self.desensitized_data.shape[1]}列")
        
        # 聚焦到输入框
        self.input_entry.focus()
        
        # 等待窗口关闭
        self.window.wait_window()
        return self.chat_history
    
    def _prepare_context(self):
        """准备数据上下文"""
        context = f"""
[系统提示：以下是脱敏后的分析数据]
数据维度: {self.desensitized_data.shape[0]}行 x {self.desensitized_data.shape[1]}列
列名: {list(self.desensitized_data.columns)}

数据样本（前3行）:
{self.desensitized_data.head(3).to_string()}

数值列统计:
{self.desensitized_data.describe().to_string() if len(self.desensitized_data.select_dtypes(include=[np.number]).columns) > 0 else '无数值列'}

请基于以上脱敏数据回答用户的问题，提供专业的数据分析建议。
"""
        return context
    
    def _display_message(self, role, content):
        """显示消息"""
        self.chat_text.config(state='normal')
        
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if role == "user":
            self.chat_text.insert(tk.END, f"[{timestamp}] 用户: ", "user")
        elif role == "assistant":
            self.chat_text.insert(tk.END, f"[{timestamp}] DeepSeek助手: ", "assistant")
        else:
            self.chat_text.insert(tk.END, f"[{timestamp}] 系统: ", "system")
        
        self.chat_text.insert(tk.END, f"{content}\n\n")
        self.chat_text.config(state='disabled')
        self.chat_text.see(tk.END)
    
    def _send_message(self):
        """发送消息"""
        user_input = self.input_var.get().strip()
        if not user_input:
            return
        
        # 清空输入框并禁用发送按钮
        self.input_var.set("")
        self.send_button.config(state="disabled")
        self.input_entry.config(state="disabled")
        
        # 显示用户消息
        self._display_message("user", user_input)
        
        # 在后台线程中调用API
        def api_call():
            try:
                response = self._call_deepseek_api(user_input)
                self.window.after(0, lambda: self._handle_response(user_input, response))
            except Exception as e:
                error_msg = str(e)
                self.window.after(0, lambda: self._handle_error(error_msg))
        
        threading.Thread(target=api_call, daemon=True).start()
    
    def _call_deepseek_api(self, user_input):
        """调用DeepSeek API - 使用OpenAI客户端"""
        if not OPENAI_AVAILABLE:
            raise Exception("OpenAI库未安装，无法使用DeepSeek API。请运行: pip install openai")
        
        # 构建消息历史
        messages = [
            {"role": "system", "content": self.context}
        ]
        
        # 添加最近的对话历史（最多10轮）
        recent_history = self.chat_history[-20:] if len(self.chat_history) > 20 else self.chat_history
        messages.extend(recent_history)
        
        # 添加当前用户输入
        messages.append({"role": "user", "content": user_input})
        
        try:
            # 创建OpenAI客户端，指向DeepSeek API
            client = OpenAI(
                api_key=self.api_key,
                base_url=self.model_config["endpoint"]
            )
            
            # 调用API
            response = client.chat.completions.create(
                model=self.model_config["model"],
                messages=messages,
                max_tokens=2000,
                temperature=self.model_config["temperature"],
                stream=self.model_config.get("stream", False)
            )
            
            if self.model_config.get("stream", False):
                # 流式响应处理
                full_content = ""
                for chunk in response:
                    if chunk.choices[0].delta.content is not None:
                        content = chunk.choices[0].delta.content
                        full_content += content
                        # 实时更新显示（可选）
                        self.window.after(0, lambda c=full_content: self._update_streaming_display(c))
                return full_content
            else:
                # 非流式响应处理
                return response.choices[0].message.content
                
        except Exception as e:
            if "OpenAI" in str(e):
                raise Exception(f"DeepSeek API调用失败: {str(e)}")
            else:
                raise Exception(f"API调用错误: {str(e)}")
    
    def _handle_response(self, user_input, ai_response):
        """处理API响应"""
        # 保存到历史
        self.chat_history.append({"role": "user", "content": user_input})
        self.chat_history.append({"role": "assistant", "content": ai_response})
        
        # 显示AI响应
        self._display_message("assistant", ai_response)
        
        # 重新启用输入
        self.send_button.config(state="normal")
        self.input_entry.config(state="normal")
        self.input_entry.focus()
    
    def _handle_error(self, error_msg):
        """处理错误"""
        self._display_message("system", f"错误: {error_msg}")
        
        # 重新启用输入
        self.send_button.config(state="normal")
        self.input_entry.config(state="normal")
        self.input_entry.focus()
    
    def _update_streaming_display(self, content):
        """更新流式显示内容"""
        # 这里可以实现实时显示流式内容的逻辑
        # 暂时不实现，避免界面闪烁
        pass
    
    def _save_chat(self):
        """保存对话历史"""
        if not self.chat_history:
            messagebox.showinfo("提示", "暂无对话记录")
            return
        
        from tkinter import filedialog
        file_path = filedialog.asksaveasfilename(
            title="保存对话记录",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                if file_path.endswith('.json'):
                    with open(file_path, 'w', encoding='utf-8') as f:
                        json.dump(self.chat_history, f, ensure_ascii=False, indent=2)
                else:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        for msg in self.chat_history:
                            role = "用户" if msg["role"] == "user" else "DeepSeek助手"
                            f.write(f"{role}: {msg['content']}\n\n")
                
                messagebox.showinfo("成功", f"对话记录已保存至: {file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {str(e)}")
    
    def _clear_chat(self):
        """清空对话"""
        if messagebox.askyesno("确认", "确定要清空所有对话记录吗？"):
            self.chat_history.clear()
            self.chat_text.config(state='normal')
            self.chat_text.delete(1.0, tk.END)
            self.chat_text.config(state='disabled')
            
            # 重新显示初始消息
            self._display_message("system", "对话已清空，DeepSeek助手重新准备就绪！")

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
        self.root.title("Excel小工具 v4.2")
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
        
        # 配置主框架的自适应
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=0)  # 标题行
        main_frame.rowconfigure(1, weight=1)  # 标签页区域
        
        # 标题
        title_label = ttk.Label(main_frame, text="Excel小工具 v4.2", 
                               font=("Arial", 18, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 15))
        
        # 创建标签页控件
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 创建三个标签页
        self.setup_compare_tab()
        self.setup_process_tab() 
        self.setup_analysis_tab()
    
    def setup_compare_tab(self):
        """设置对比工具标签页"""
        # 创建对比工具标签页框架
        compare_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(compare_frame, text="对比工具")
        
        # 配置对比工具框架的自适应 - 每列都有相等权重，可以灵活调整
        compare_frame.columnconfigure(0, weight=1)
        compare_frame.columnconfigure(1, weight=1)
        
        # 功能说明
        desc_text = """功能特点：
• 基于指定列进行精确行匹配    • 智能选择参与对比的列，关键列必选
• 智能处理数据类型差异        • 竖排排列，差异数据红色粗体高亮显示
• 弹出窗口选择比对列         • 自动生成序号的差异报告
• 并排布局优化界面           • 精简操作流程"""
        desc_label = ttk.Label(compare_frame, text=desc_text, font=("Arial", 9), 
                              foreground="darkblue")
        desc_label.grid(row=0, column=0, columnspan=2, pady=(0, 15))
        
        # 第一行：主表格选择 | 关键列选择（并排）
        # 左侧：主表格选择
        main_file_frame = ttk.LabelFrame(compare_frame, text="1. 选择主对比表格", padding="10")
        main_file_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10), pady=(0, 15))
        main_file_frame.columnconfigure(0, weight=1)
        
        self.main_file_var = tk.StringVar()
        ttk.Entry(main_file_frame, textvariable=self.main_file_var, 
                 state="readonly", font=("Arial", 10)).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10), pady=(0, 10))
        ttk.Button(main_file_frame, text="选择文件", 
                  command=self.select_main_file).grid(row=0, column=1, pady=(0, 10))
        
        # 右侧：关键列选择
        key_column_frame = ttk.LabelFrame(compare_frame, text="2. 选择用于行匹配的关键列", padding="10")
        key_column_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
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
        files_frame = ttk.LabelFrame(compare_frame, text="3. 添加待对比文件", padding="10")
        files_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10), pady=(0, 15))
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
        settings_frame = ttk.LabelFrame(compare_frame, text="4. 比对设置", padding="10")
        settings_frame.grid(row=2, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
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
        self.exclude_checkboxes_frame = ttk.Frame(compare_frame)
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
        action_frame = ttk.Frame(compare_frame)
        action_frame.grid(row=3, column=0, columnspan=2, pady=(0, 15))
        
        ttk.Label(action_frame, text="5. 执行对比:", 
                 font=("Arial", 11, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        
        self.compare_button = ttk.Button(action_frame, text="开始对比", 
                                        command=self.start_compare, 
                                        style="Accent.TButton")
        self.compare_button.grid(row=0, column=1, padx=(0, 10))
        
        ttk.Button(action_frame, text="重置", 
                  command=self.reset).grid(row=0, column=2)
        
        # 6. 日志区域
        ttk.Label(compare_frame, text="6. 处理日志:", 
                 font=("Arial", 11, "bold")).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        log_frame = ttk.Frame(compare_frame)
        log_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, font=("Consolas", 9))
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(compare_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 9))
        status_bar.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 完善自适应配置 - 设置行列权重以确保各区域能够自适应调整
        # 行权重配置：让需要伸缩的区域有合适的权重
        compare_frame.rowconfigure(0, weight=0)  # 功能说明行，固定高度
        compare_frame.rowconfigure(1, weight=0)  # 主表格和关键列选择行，最小高度
        compare_frame.rowconfigure(2, weight=3)  # 文件列表和设置区域，主要伸缩区域，给予最大权重
        compare_frame.rowconfigure(3, weight=0)  # 执行对比按钮行，固定高度
        compare_frame.rowconfigure(4, weight=0)  # 日志标题行，固定高度
        compare_frame.rowconfigure(5, weight=2)  # 日志区域，次要伸缩区域
        compare_frame.rowconfigure(6, weight=0)  # 状态栏，固定高度
        
        # 初始化日志
        self.log("Excel小工具 v4.2 已启动")
        self.log("新功能：标签页界面，支持多种Excel处理工具")
        self.log("对比工具：比对列选择，弹出窗口智能选择，关键列必选机制")
        self.log("界面优化：主表格和关键列并排，待对比文件和比对设置并排")
    
    def setup_process_tab(self):
        """设置表格处理标签页"""
        # 创建表格处理标签页框架
        process_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(process_frame, text="表格处理")
        
        # 配置框架
        process_frame.columnconfigure(0, weight=1)
        process_frame.rowconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(process_frame, text="表格处理工具", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # 功能说明
        desc_text = """即将推出的功能：
• 批量格式转换（xlsx ↔ csv）     • 工作表合并与拆分
• 数据清洗与去重               • 格式标准化处理
• 列操作与重组                 • 批量文件处理
        
🚧 功能开发中，敬请期待..."""
        
        desc_label = ttk.Label(process_frame, text=desc_text, 
                              font=("Arial", 11), justify=tk.CENTER,
                              foreground="gray")
        desc_label.grid(row=1, column=0, pady=50)
    
    def setup_analysis_tab(self):
        """设置数据分析标签页"""
        # 创建数据分析标签页框架
        analysis_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(analysis_frame, text="数据分析")
        
        # 配置框架的自适应
        analysis_frame.columnconfigure(0, weight=1)
        analysis_frame.columnconfigure(1, weight=1)
        
        # 功能说明
        desc_text = """功能特点：
• 多文件Excel数据合并分析      • 智能数据脱敏保护隐私
• 敏感列和敏感行双重保护       • DeepSeek对话式分析
• 交互式智能问答              • 综合分析报告生成
• 对话历史完整记录            • 支持文件来源追踪"""
        desc_label = ttk.Label(analysis_frame, text=desc_text, font=("Arial", 9), 
                              foreground="darkblue")
        desc_label.grid(row=0, column=0, columnspan=2, pady=(0, 15))
        
        # 第一行：数据导入（全宽）
        import_frame = ttk.LabelFrame(analysis_frame, text="1. 数据导入", padding="10")
        import_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        import_frame.columnconfigure(0, weight=1)
        import_frame.rowconfigure(0, weight=1)
        
        # 文件列表显示区域
        files_list_frame = ttk.Frame(import_frame)
        files_list_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        files_list_frame.columnconfigure(0, weight=1)
        files_list_frame.rowconfigure(0, weight=1)
        
        # 创建Treeview来显示文件列表
        self.analysis_files_tree = ttk.Treeview(files_list_frame, columns=('file', 'status'), show='headings', height=4)
        self.analysis_files_tree.heading('file', text='Excel文件列表')
        self.analysis_files_tree.heading('status', text='状态')
        self.analysis_files_tree.column('file', width=300)
        self.analysis_files_tree.column('status', width=100)
        self.analysis_files_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 滚动条
        analysis_files_scrollbar = ttk.Scrollbar(files_list_frame, orient=tk.VERTICAL, command=self.analysis_files_tree.yview)
        analysis_files_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.analysis_files_tree.configure(yscrollcommand=analysis_files_scrollbar.set)
        
        # 文件操作按钮
        files_buttons_frame = ttk.Frame(import_frame)
        files_buttons_frame.grid(row=0, column=1, sticky=(tk.N))
        
        ttk.Button(files_buttons_frame, text="添加Excel文件", 
                  command=self.select_analysis_file).grid(row=0, column=0, pady=(0, 5))
        ttk.Button(files_buttons_frame, text="移除选中", 
                  command=self.remove_analysis_file).grid(row=1, column=0, pady=(0, 5))
        ttk.Button(files_buttons_frame, text="清空列表", 
                  command=self.clear_analysis_files).grid(row=2, column=0)
        
        # 第二行：敏感数据配置（三列布局：列选择 | 行选择 | 分析设置）
        # 配置三列权重
        analysis_frame.columnconfigure(0, weight=1)  # 列选择
        analysis_frame.columnconfigure(1, weight=1)  # 行选择  
        analysis_frame.columnconfigure(2, weight=1)  # 分析设置
        
        # 左侧：敏感列配置
        sensitive_col_frame = ttk.LabelFrame(analysis_frame, text="2. 敏感列配置", padding="10")
        sensitive_col_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5), pady=(0, 15))
        sensitive_col_frame.columnconfigure(0, weight=1)
        sensitive_col_frame.rowconfigure(1, weight=1)
        
        ttk.Label(sensitive_col_frame, text="选择需要脱敏的列：", font=("Arial", 9)).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # 敏感列选择列表
        sensitive_col_list_frame = ttk.Frame(sensitive_col_frame)
        sensitive_col_list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        sensitive_col_list_frame.columnconfigure(0, weight=1)
        sensitive_col_list_frame.rowconfigure(0, weight=1)
        
        self.sensitive_listbox = tk.Listbox(sensitive_col_list_frame, selectmode=tk.MULTIPLE, height=6)
        self.sensitive_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        sensitive_col_scrollbar = ttk.Scrollbar(sensitive_col_list_frame, orient=tk.VERTICAL, command=self.sensitive_listbox.yview)
        sensitive_col_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.sensitive_listbox.configure(yscrollcommand=sensitive_col_scrollbar.set)
        
        # 中间：敏感行配置
        sensitive_row_frame = ttk.LabelFrame(analysis_frame, text="3. 敏感行配置", padding="10")
        sensitive_row_frame.grid(row=2, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 5), pady=(0, 15))
        sensitive_row_frame.columnconfigure(0, weight=1)
        sensitive_row_frame.rowconfigure(2, weight=1)
        
        ttk.Label(sensitive_row_frame, text="选择需要脱敏的行：", font=("Arial", 9)).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # 行选择方式
        self.row_select_mode = tk.StringVar(value="none")
        ttk.Radiobutton(sensitive_row_frame, text="不选择行", variable=self.row_select_mode, 
                       value="none").grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        
        # 行选择列表
        sensitive_row_list_frame = ttk.Frame(sensitive_row_frame)
        sensitive_row_list_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        sensitive_row_list_frame.columnconfigure(0, weight=1)
        sensitive_row_list_frame.rowconfigure(0, weight=1)
        
        self.sensitive_row_listbox = tk.Listbox(sensitive_row_list_frame, selectmode=tk.MULTIPLE, height=4)
        self.sensitive_row_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        sensitive_row_scrollbar = ttk.Scrollbar(sensitive_row_list_frame, orient=tk.VERTICAL, command=self.sensitive_row_listbox.yview)
        sensitive_row_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.sensitive_row_listbox.configure(yscrollcommand=sensitive_row_scrollbar.set)
        
        # 右侧：DeepSeek对话设置
        dialog_frame = ttk.LabelFrame(analysis_frame, text="4. DeepSeek对话分析", padding="10")
        dialog_frame.grid(row=2, column=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0), pady=(0, 15))
        dialog_frame.columnconfigure(0, weight=1)
        dialog_frame.rowconfigure(1, weight=1)
        
        # 对话区域
        ttk.Label(dialog_frame, text="快速对话：", font=("Arial", 9)).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # 对话显示区域
        dialog_display_frame = ttk.Frame(dialog_frame)
        dialog_display_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        dialog_display_frame.columnconfigure(0, weight=1)
        dialog_display_frame.rowconfigure(0, weight=1)
        
        self.dialog_text = scrolledtext.ScrolledText(dialog_display_frame, height=5, font=("Arial", 9), wrap=tk.WORD)
        self.dialog_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 输入和发送
        input_frame = ttk.Frame(dialog_frame)
        input_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(0, weight=1)
        
        self.dialog_input_var = tk.StringVar()
        self.dialog_input = ttk.Entry(input_frame, textvariable=self.dialog_input_var, font=("Arial", 9))
        self.dialog_input.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        self.dialog_input.bind('<Return>', self.send_dialog_message)
        
        ttk.Button(input_frame, text="发送", command=self.send_dialog_message).grid(row=0, column=1)
        
        # 高级对话按钮
        ttk.Button(dialog_frame, text="打开高级对话窗口", 
                  command=self.open_advanced_chat, 
                  style="Accent.TButton").grid(row=3, column=0, pady=(10, 0))
        
        # 输出目录
        ttk.Label(dialog_frame, text="输出目录：", font=("Arial", 9)).grid(row=4, column=0, sticky=tk.W, pady=(10, 5))
        output_controls_frame = ttk.Frame(dialog_frame)
        output_controls_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        output_controls_frame.columnconfigure(0, weight=1)
        
        self.analysis_output_var = tk.StringVar()
        ttk.Entry(output_controls_frame, textvariable=self.analysis_output_var, 
                 state="readonly", font=("Arial", 9)).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(output_controls_frame, text="选择", 
                  command=self.select_analysis_output).grid(row=0, column=1)
        
        # 第三行：生成报告
        action_frame = ttk.Frame(analysis_frame)
        action_frame.grid(row=3, column=0, columnspan=3, pady=(0, 15))
        
        ttk.Label(action_frame, text="5. 生成报告:", 
                 font=("Arial", 11, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        
        self.analysis_button = ttk.Button(action_frame, text="生成报告", 
                                        command=self.start_analysis, 
                                        style="Accent.TButton")
        self.analysis_button.grid(row=0, column=1, padx=(0, 10))
        
        ttk.Button(action_frame, text="重置", 
                  command=self.reset_analysis).grid(row=0, column=2)
        
        # 第四行：分析日志
        ttk.Label(analysis_frame, text="6. 分析日志:", 
                 font=("Arial", 11, "bold")).grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        log_frame = ttk.Frame(analysis_frame)
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.analysis_log_text = scrolledtext.ScrolledText(log_frame, height=6, font=("Consolas", 9))
        self.analysis_log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 状态栏
        self.analysis_status_var = tk.StringVar(value="就绪")
        analysis_status_bar = ttk.Label(analysis_frame, textvariable=self.analysis_status_var, 
                                      relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 9))
        analysis_status_bar.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 配置行权重
        analysis_frame.rowconfigure(0, weight=0)  # 功能说明
        analysis_frame.rowconfigure(1, weight=1)  # 数据导入和API配置
        analysis_frame.rowconfigure(2, weight=2)  # 敏感数据和分析设置
        analysis_frame.rowconfigure(3, weight=0)  # 执行按钮
        analysis_frame.rowconfigure(4, weight=0)  # 日志标题
        analysis_frame.rowconfigure(5, weight=1)  # 日志区域
        analysis_frame.rowconfigure(6, weight=0)  # 状态栏
        
        # 初始化变量
        self.analysis_files = []  # 存储选择的文件路径列表
        self.analysis_data = None
        self.analysis_columns = []
        self.dialog_history = []  # 对话历史
        self.base_url_var = tk.StringVar(value="https://api.deepseek.com")  # 添加base_url变量
        
        # 初始化日志
        self.analysis_log("数据分析工具已启动")
        self.analysis_log("支持多文件Excel数据合并、脱敏、DeepSeek对话分析和综合报告生成")
        
        # 初始化对话
        self.dialog_text.insert(tk.END, "DeepSeek助手: 您好！我是您的数据分析助手。请先上传Excel文件，然后我们可以开始分析对话。\n\n")
        self.dialog_text.config(state=tk.DISABLED)
    
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
        self.log("Excel表格对比工具 - 改进版 v4.2 已重置")
        self.log("新功能：比对列选择，弹出窗口智能选择，关键列必选机制")
        self.log("新特性：差异报告自动序号命名，只对比选中的列")
        self.log("界面优化：主表格和关键列并排，待对比文件和比对设置并排")
        self.log("✓ 所有输入和状态已重置到初始状态")
        
        # 更新状态栏
        self.update_status("就绪 - 已重置所有状态")
    
    # ==================== 数据分析功能方法 ====================
    
    def analysis_log(self, message):
        """添加分析日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.analysis_log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.analysis_log_text.see(tk.END)
        self.root.update_idletasks()
    
    def analysis_update_status(self, message):
        """更新分析状态栏"""
        self.analysis_status_var.set(message)
        self.root.update_idletasks()
    
    def select_analysis_file(self):
        """选择分析文件（支持多文件选择）"""
        file_paths = filedialog.askopenfilenames(
            title="选择Excel文件进行分析（可多选）",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        for file_path in file_paths:
            if file_path not in self.analysis_files:
                try:
                    # 验证文件是否可以读取
                    test_data = pd.read_excel(file_path)
                    
                    # 添加到文件列表
                    self.analysis_files.append(file_path)
                    
                    # 添加到树形视图
                    self.analysis_files_tree.insert('', tk.END, values=(os.path.basename(file_path), "已加载"))
                    
                    self.analysis_log(f"✓ 添加文件: {os.path.basename(file_path)}")
                    self.analysis_log(f"  数据维度: {test_data.shape[0]}行 x {test_data.shape[1]}列")
                    
                except Exception as e:
                    self.analysis_log(f"✗ 文件加载失败: {os.path.basename(file_path)} - {str(e)}")
                    continue
        
        # 如果有文件被添加，合并数据并更新界面
        if self.analysis_files:
            self._merge_analysis_data()
            self.analysis_update_status(f"已加载 {len(self.analysis_files)} 个文件")
    
    def _merge_analysis_data(self):
        """合并多个Excel文件的数据"""
        try:
            all_dataframes = []
            
            for file_path in self.analysis_files:
                df = pd.read_excel(file_path)
                # 添加文件来源列
                df['数据来源文件'] = os.path.basename(file_path)
                all_dataframes.append(df)
            
            if all_dataframes:
                # 合并所有数据框
                self.analysis_data = pd.concat(all_dataframes, ignore_index=True, sort=False)
                self.analysis_columns = list(self.analysis_data.columns)
                
                self.analysis_log(f"✓ 数据合并完成，总计: {self.analysis_data.shape[0]}行 x {self.analysis_data.shape[1]}列")
                
                # 更新敏感列列表
                self.sensitive_listbox.delete(0, tk.END)
                for col in self.analysis_columns:
                    self.sensitive_listbox.insert(tk.END, col)
                
                # 更新敏感行列表（只显示前100行，避免界面卡顿）
                self.sensitive_row_listbox.delete(0, tk.END)
                max_rows = min(100, len(self.analysis_data))
                for i in range(max_rows):
                    # 显示行号和第一列的值作为标识
                    first_col_value = str(self.analysis_data.iloc[i, 0]) if len(self.analysis_data.columns) > 0 else ""
                    row_label = f"第{i+1}行: {first_col_value[:20]}{'...' if len(first_col_value) > 20 else ''}"
                    self.sensitive_row_listbox.insert(tk.END, row_label)
                
                if len(self.analysis_data) > 100:
                    self.sensitive_row_listbox.insert(tk.END, f"... 还有 {len(self.analysis_data) - 100} 行数据")
                
                # 更新对话区域
                self.dialog_text.config(state=tk.NORMAL)
                self.dialog_text.insert(tk.END, f"系统: 已合并 {len(self.analysis_files)} 个Excel文件，总计 {self.analysis_data.shape[0]} 行 {self.analysis_data.shape[1]} 列数据。您可以开始提问了！\n\n")
                self.dialog_text.config(state=tk.DISABLED)
                self.dialog_text.see(tk.END)
                
        except Exception as e:
            self.analysis_log(f"✗ 数据合并失败: {str(e)}")
            messagebox.showerror("错误", f"数据合并失败: {str(e)}")
    
    def remove_analysis_file(self):
        """移除选中的分析文件"""
        selected_items = self.analysis_files_tree.selection()
        if not selected_items:
            messagebox.showwarning("提示", "请先选择要移除的文件")
            return
        
        for item in selected_items:
            # 获取文件名
            values = self.analysis_files_tree.item(item, 'values')
            if values:
                filename = values[0]
                # 从列表中移除
                self.analysis_files = [f for f in self.analysis_files if os.path.basename(f) != filename]
                # 从树形视图中移除
                self.analysis_files_tree.delete(item)
                self.analysis_log(f"✓ 移除文件: {filename}")
        
        # 重新合并数据
        if self.analysis_files:
            self._merge_analysis_data()
            self.analysis_update_status(f"当前有 {len(self.analysis_files)} 个文件")
        else:
            # 清空数据
            self.analysis_data = None
            self.analysis_columns = []
            self.sensitive_listbox.delete(0, tk.END)
            self.sensitive_row_listbox.delete(0, tk.END)
            
            # 更新对话区域
            self.dialog_text.config(state=tk.NORMAL)
            self.dialog_text.insert(tk.END, "系统: 所有文件已移除，请重新选择Excel文件。\n\n")
            self.dialog_text.config(state=tk.DISABLED)
            self.dialog_text.see(tk.END)
            
            self.analysis_update_status("无文件")
    
    def clear_analysis_files(self):
        """清空所有分析文件"""
        if self.analysis_files:
            result = show_confirm_dialog("确认清空", "确定要清空所有Excel文件吗？")
            if result:
                self.analysis_files.clear()
                self.analysis_files_tree.delete(*self.analysis_files_tree.get_children())
                
                # 清空数据
                self.analysis_data = None
                self.analysis_columns = []
                self.sensitive_listbox.delete(0, tk.END)
                self.sensitive_row_listbox.delete(0, tk.END)
                
                # 更新对话区域
                self.dialog_text.config(state=tk.NORMAL)
                self.dialog_text.insert(tk.END, "系统: 所有文件已清空，请重新选择Excel文件。\n\n")
                self.dialog_text.config(state=tk.DISABLED)
                self.dialog_text.see(tk.END)
                
                self.analysis_log("✓ 已清空所有Excel文件")
                self.analysis_update_status("文件列表已清空")
    
    def send_dialog_message(self, event=None):
        """发送对话消息"""
        message = self.dialog_input_var.get().strip()
        if not message:
            return
        
        # 清空输入框
        self.dialog_input_var.set("")
        
        # 显示用户消息
        self.dialog_text.config(state=tk.NORMAL)
        self.dialog_text.insert(tk.END, f"用户: {message}\n")
        self.dialog_text.config(state=tk.DISABLED)
        self.dialog_text.see(tk.END)
        
        # 检查是否有数据
        if self.analysis_data is None:
            self.dialog_text.config(state=tk.NORMAL)
            self.dialog_text.insert(tk.END, "DeepSeek助手: 请先上传Excel文件，然后我们可以开始分析。\n\n")
            self.dialog_text.config(state=tk.DISABLED)
            self.dialog_text.see(tk.END)
            return
        
        # 检查API配置
        api_key = ENV_CONFIG['api_key']
        if not api_key:
            self.dialog_text.config(state=tk.NORMAL)
            self.dialog_text.insert(tk.END, "DeepSeek助手: 请先在.env文件中配置API Key才能进行智能分析。\n\n")
            self.dialog_text.config(state=tk.DISABLED)
            self.dialog_text.see(tk.END)
            return
        
        # 在后台线程中调用API
        def api_call():
            try:
                # 准备数据摘要
                data_summary = self.prepare_data_summary()
                
                # 构建对话历史
                messages = [
                    {"role": "system", "content": f"你是一个专业的数据分析师。当前分析的数据概况：{data_summary}"},
                ]
                
                # 添加历史对话
                for hist in self.dialog_history[-5:]:  # 只保留最近5轮对话
                    messages.append(hist)
                
                # 添加当前用户消息
                messages.append({"role": "user", "content": message})
                
                # 调用API
                response = self.call_deepseek_api(messages)
                
                # 更新界面
                self.root.after(0, lambda: self.update_dialog_response(message, response))
                
            except Exception as e:
                error_msg = str(e)
                self.root.after(0, lambda: self.update_dialog_error(error_msg))
        
        threading.Thread(target=api_call, daemon=True).start()
    
    def prepare_data_summary(self):
        """准备数据摘要"""
        if self.analysis_data is None:
            return "无数据"
        
        # 获取敏感列
        selected_col_indices = self.sensitive_listbox.curselection()
        sensitive_columns = [self.analysis_columns[i] for i in selected_col_indices]
        
        # 获取敏感行
        selected_row_indices = self.sensitive_row_listbox.curselection()
        
        # 创建脱敏数据用于分析
        clean_data = self.analysis_data.copy()
        for col in sensitive_columns:
            clean_data[col] = f"[脱敏_{col}]"
        
        summary = {
            "数据维度": f"{clean_data.shape[0]}行 x {clean_data.shape[1]}列",
            "列名": list(clean_data.columns),
            "脱敏列": sensitive_columns,
            "脱敏行数": len(selected_row_indices),
            "数值列统计": clean_data.describe().to_dict() if len(clean_data.select_dtypes(include=[np.number]).columns) > 0 else "无数值列",
            "数据样本": clean_data.head(3).to_dict() if len(clean_data) > 0 else "无数据"
        }
        
        return str(summary)
    
    def call_deepseek_api(self, messages):
        """调用DeepSeek API - 使用OpenAI客户端"""
        if not OPENAI_AVAILABLE:
            raise Exception("OpenAI库未安装，无法使用DeepSeek API。请运行: pip install openai")
        
        api_key = ENV_CONFIG['api_key']
        base_url = ENV_CONFIG['base_url']
        
        try:
            # 创建OpenAI客户端，指向DeepSeek API
            client = OpenAI(
                api_key=api_key,
                base_url=base_url
            )
            
            # 调用API
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=messages,
                max_tokens=1500,
                temperature=0.7
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            if "OpenAI" in str(e):
                raise Exception(f"DeepSeek API调用失败: {str(e)}")
            else:
                raise Exception(f"API调用错误: {str(e)}")
    
    def update_dialog_response(self, user_message, ai_response):
        """更新对话响应"""
        # 保存到历史
        self.dialog_history.append({"role": "user", "content": user_message})
        self.dialog_history.append({"role": "assistant", "content": ai_response})
        
        # 显示AI响应
        self.dialog_text.config(state=tk.NORMAL)
        self.dialog_text.insert(tk.END, f"DeepSeek助手: {ai_response}\n\n")
        self.dialog_text.config(state=tk.DISABLED)
        self.dialog_text.see(tk.END)
    
    def update_dialog_error(self, error_msg):
        """更新对话错误"""
        self.dialog_text.config(state=tk.NORMAL)
        self.dialog_text.insert(tk.END, f"系统错误: {error_msg}\n\n")
        self.dialog_text.config(state=tk.DISABLED)
        self.dialog_text.see(tk.END)
    
    def open_advanced_chat(self):
        """打开高级对话窗口"""
        # 检查是否有数据
        if self.analysis_data is None:
            messagebox.showwarning("提示", "请先上传Excel文件")
            return
        
        # 检查API配置
        api_key = ENV_CONFIG['api_key']
        if not api_key:
            messagebox.showwarning("提示", "请先在.env文件中配置API Key")
            return
        
        # 显示模型选择窗口
        model_selector = ModelSelectionWindow(self.root, api_key)
        model_config = model_selector.show()
        
        if model_config is None:
            return  # 用户取消了选择
        
        # 准备脱敏数据
        selected_col_indices = self.sensitive_listbox.curselection()
        sensitive_columns = [self.analysis_columns[i] for i in selected_col_indices]
        
        selected_row_indices = self.sensitive_row_listbox.curselection()
        
        # 创建脱敏数据
        clean_data = self.analysis_data.copy()
        for col in sensitive_columns:
            clean_data[col] = f"[脱敏_{col}]"
        
        # 行脱敏
        if selected_row_indices:
            for row_idx in selected_row_indices:
                for col in clean_data.columns:
                    if col not in sensitive_columns:
                        clean_data.iloc[row_idx, clean_data.columns.get_loc(col)] = f"[脱敏_行{row_idx+1}]"
        
        # 打开对话窗口
        chat_window = ChatWindow(self.root, clean_data, api_key, model_config)
        chat_history = chat_window.show()
        
        # 将对话历史合并到主对话历史中
        if chat_history:
            self.dialog_history.extend(chat_history)
            
            # 在快速对话区域显示摘要
            self.dialog_text.config(state=tk.NORMAL)
            self.dialog_text.insert(tk.END, f"\n=== 高级对话完成 ===\n")
            self.dialog_text.insert(tk.END, f"模型: {model_config['model']}\n")
            self.dialog_text.insert(tk.END, f"对话轮数: {len(chat_history) // 2}\n")
            self.dialog_text.insert(tk.END, f"已合并到分析历史中\n\n")
            self.dialog_text.config(state=tk.DISABLED)
            self.dialog_text.see(tk.END)
            
            self.analysis_log(f"✓ 高级对话完成，模型: {model_config['model']}，对话轮数: {len(chat_history) // 2}")
    
    def select_analysis_output(self):
        """选择分析输出目录"""
        directory = filedialog.askdirectory(title="选择分析结果输出目录")
        if directory:
            self.analysis_output_var.set(directory)
            self.analysis_log(f"✓ 输出目录: {directory}")
            self.analysis_update_status("输出目录已选择")
    
    def start_analysis(self):
        """开始数据分析"""
        # 验证输入
        if self.analysis_data is None:
            messagebox.showwarning("提示", "请先选择Excel文件")
            return
        
        if not self.analysis_output_var.get():
            messagebox.showwarning("提示", "请先选择输出目录")
            return
        
        # 禁用分析按钮
        self.analysis_button.config(state="disabled")
        self.analysis_log("开始执行数据分析...")
        self.analysis_update_status("正在分析数据...")
        
        def analysis_thread():
            try:
                # 获取敏感列
                selected_col_indices = self.sensitive_listbox.curselection()
                sensitive_columns = [self.analysis_columns[i] for i in selected_col_indices]
                
                # 获取敏感行
                selected_row_indices = self.sensitive_row_listbox.curselection()
                
                # 数据脱敏
                clean_data = self.analysis_data.copy()
                for col in sensitive_columns:
                    clean_data[col] = f"[脱敏_{col}]"
                
                # 行脱敏（如果选择了敏感行）
                if selected_row_indices:
                    for row_idx in selected_row_indices:
                        for col in clean_data.columns:
                            if col not in sensitive_columns:  # 避免重复脱敏
                                clean_data.iloc[row_idx, clean_data.columns.get_loc(col)] = f"[脱敏_行{row_idx+1}]"
                
                self.analysis_log(f"✓ 数据脱敏完成，脱敏列: {sensitive_columns}，脱敏行: {len(selected_row_indices)}行")
                
                # 生成图表
                chart_path = self.generate_chart(clean_data, sensitive_columns)
                
                # 整合对话历史作为AI分析
                dialog_summary = self.generate_dialog_summary()
                
                # 生成最终报告
                report_path = self.generate_final_report(clean_data, chart_path, dialog_summary, sensitive_columns, selected_row_indices)
                
                self.root.after(0, lambda: self.analysis_complete(True, report_path))
                
            except Exception as e:
                self.root.after(0, lambda: self.analysis_complete(False, str(e)))
        
        threading.Thread(target=analysis_thread, daemon=True).start()
    
    def generate_dialog_summary(self):
        """生成对话摘要"""
        if not self.dialog_history:
            return "未进行DeepSeek对话分析"
        
        summary = "=== DeepSeek对话分析完整记录 ===\n\n"
        summary += f"总对话轮数: {len(self.dialog_history) // 2}\n"
        summary += f"分析时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        
        # 整理对话内容
        current_round = 1
        for i in range(0, len(self.dialog_history), 2):
            if i + 1 < len(self.dialog_history):
                user_msg = self.dialog_history[i]
                assistant_msg = self.dialog_history[i + 1]
                
                summary += f"--- 第{current_round}轮对话 ---\n"
                summary += f"用户: {user_msg['content']}\n\n"
                summary += f"DeepSeek助手: {assistant_msg['content']}\n\n"
                current_round += 1
        
        summary += "=== 对话分析记录结束 ==="
        return summary
    
    def generate_chart(self, data, sensitive_columns):
        """生成数据图表"""
        try:
            import matplotlib.pyplot as plt
            import matplotlib
            matplotlib.use('Agg')  # 使用非交互式后端
            
            # 设置中文字体
            plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
            plt.rcParams['axes.unicode_minus'] = False
            
            # 选择数值列进行分析
            numeric_columns = data.select_dtypes(include=[np.number]).columns.tolist()
            # 排除敏感列
            numeric_columns = [col for col in numeric_columns if col not in sensitive_columns]
            
            if not numeric_columns:
                self.analysis_log("⚠ 没有可用的数值列进行图表生成")
                return None
            
            chart_type = self.chart_type_var.get()
            
            plt.figure(figsize=(10, 6))
            
            if chart_type == "柱状图":
                if len(numeric_columns) >= 1:
                    col = numeric_columns[0]
                    data[col].value_counts().head(10).plot(kind='bar')
                    plt.title(f'{col} - 柱状图')
                    plt.ylabel('频次')
            
            elif chart_type == "折线图":
                if len(numeric_columns) >= 1:
                    col = numeric_columns[0]
                    data[col].plot(kind='line')
                    plt.title(f'{col} - 折线图')
                    plt.ylabel(col)
            
            elif chart_type == "饼图":
                if len(numeric_columns) >= 1:
                    col = numeric_columns[0]
                    data[col].value_counts().head(5).plot(kind='pie', autopct='%1.1f%%')
                    plt.title(f'{col} - 饼图')
            
            elif chart_type == "散点图":
                if len(numeric_columns) >= 2:
                    plt.scatter(data[numeric_columns[0]], data[numeric_columns[1]])
                    plt.xlabel(numeric_columns[0])
                    plt.ylabel(numeric_columns[1])
                    plt.title(f'{numeric_columns[0]} vs {numeric_columns[1]} - 散点图')
            
            elif chart_type == "箱线图":
                if len(numeric_columns) >= 1:
                    data[numeric_columns[:3]].boxplot()
                    plt.title('数值列箱线图')
            
            plt.tight_layout()
            
            # 保存图表
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            chart_filename = f"数据分析图表_{timestamp}.png"
            chart_path = os.path.join(self.analysis_output_var.get(), chart_filename)
            plt.savefig(chart_path, dpi=300, bbox_inches='tight')
            plt.close()
            
            self.analysis_log(f"✓ 图表生成完成: {chart_filename}")
            return chart_path
            
        except Exception as e:
            self.analysis_log(f"✗ 图表生成失败: {str(e)}")
            return None
    

    
    def generate_final_report(self, data, chart_path, dialog_summary, sensitive_columns, sensitive_rows):
        """生成最终分析报告"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_filename = f"数据分析报告_{timestamp}.xlsx"
            report_path = os.path.join(self.analysis_output_var.get(), report_filename)
            
            with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
                # 写入脱敏数据
                data.to_excel(writer, sheet_name='脱敏数据', index=False)
                
                # 写入统计摘要
                if len(data.select_dtypes(include=[np.number]).columns) > 0:
                    data.describe().to_excel(writer, sheet_name='统计摘要')
                
                # 写入对话分析报告
                if dialog_summary:
                    report_df = pd.DataFrame({'DeepSeek对话分析': [dialog_summary]})
                    report_df.to_excel(writer, sheet_name='对话分析报告', index=False)
                
                # 写入元数据
                metadata = {
                    '分析时间': [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                    '数据维度': [f"{data.shape[0]}行 x {data.shape[1]}列"],
                    '脱敏列': [', '.join(sensitive_columns) if sensitive_columns else '无'],
                    '脱敏行数': [len(sensitive_rows)],
                    '图表文件': [os.path.basename(chart_path) if chart_path else '无'],
                    '对话轮数': [len(self.dialog_history) // 2]
                }
                metadata_df = pd.DataFrame(metadata)
                metadata_df.to_excel(writer, sheet_name='分析元数据', index=False)
            
            self.analysis_log(f"✓ 分析报告生成完成: {report_filename}")
            return report_path
            
        except Exception as e:
            self.analysis_log(f"✗ 报告生成失败: {str(e)}")
            return None
    
    def analysis_complete(self, success, result):
        """分析完成回调"""
        self.analysis_button.config(state="normal")
        
        if success:
            self.analysis_log("✓ 数据分析完成！")
            self.analysis_update_status("分析完成")
            messagebox.showinfo("成功", f"数据分析完成！\n报告已保存至: {result}")
            
            # 自动打开输出目录
            if os.path.exists(result):
                if open_and_highlight_file(result):
                    self.analysis_log("✓ 已自动打开分析报告保存目录")
                else:
                    self.analysis_log("✗ 打开目录失败，请手动查看")
        else:
            self.analysis_log(f"✗ 分析失败: {result}")
            self.analysis_update_status("分析失败")
            messagebox.showerror("错误", f"数据分析失败: {result}")
    
    def reset_analysis(self):
        """重置分析设置"""
        result = show_confirm_dialog("确认重置", "确定要重置所有分析设置吗？")
        if not result:
            return
        
        # 重置所有变量
        self.analysis_files.clear()
        self.base_url_var.set("https://api.deepseek.com")
        self.analysis_output_var.set("")
        self.dialog_input_var.set("")
        self.row_select_mode.set("none")
        
        # 清空控件
        self.analysis_files_tree.delete(*self.analysis_files_tree.get_children())
        self.sensitive_listbox.delete(0, tk.END)
        self.sensitive_row_listbox.delete(0, tk.END)
        self.analysis_log_text.delete(1.0, tk.END)
        
        # 重置对话
        self.dialog_text.config(state=tk.NORMAL)
        self.dialog_text.delete(1.0, tk.END)
        self.dialog_text.insert(tk.END, "DeepSeek助手: 您好！我是您的数据分析助手。请先上传Excel文件，然后我们可以开始分析对话。\n\n")
        self.dialog_text.config(state=tk.DISABLED)
        
        # 重置数据
        self.analysis_data = None
        self.analysis_columns = []
        self.dialog_history = []
        
        # 重新初始化日志
        self.analysis_log("数据分析工具已重置")
        self.analysis_log("支持多文件Excel数据脱敏、DeepSeek对话分析和综合报告生成")
        self.analysis_update_status("就绪 - 已重置所有设置")


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