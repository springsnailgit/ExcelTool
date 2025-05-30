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
    dialog.resizable(True, True)  # 允许调整大小
    dialog.minsize(400, 200)  # 设置最小尺寸
    dialog.transient()
    dialog.grab_set()
    
    # 配置对话框的自适应
    dialog.columnconfigure(0, weight=1)
    dialog.rowconfigure(0, weight=1)
    dialog.rowconfigure(1, weight=0)
    
    result = tk.BooleanVar()
    result.set(False)
    
    # 主框架
    main_frame = ttk.Frame(dialog, padding="20")
    main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    main_frame.columnconfigure(0, weight=1)
    main_frame.rowconfigure(0, weight=1)
    main_frame.rowconfigure(1, weight=0)
    
    # 消息文本 - 使用Text组件以更好地显示多行内容
    text_frame = ttk.Frame(main_frame)
    text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
    text_frame.columnconfigure(0, weight=1)
    text_frame.rowconfigure(0, weight=1)
    
    text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Arial", 10),
                         height=6, width=50, relief=tk.FLAT,
                         background=dialog.cget('bg'), state=tk.DISABLED)
    text_widget.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    # 添加滚动条
    scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
    scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    text_widget.configure(yscrollcommand=scrollbar.set)
    
    # 插入消息内容
    text_widget.config(state=tk.NORMAL)
    text_widget.insert(tk.END, message)
    text_widget.config(state=tk.DISABLED)
    
    # 按钮框架 - 居中对称布局
    button_frame = ttk.Frame(main_frame)
    button_frame.grid(row=1, column=0, pady=(10, 0))
    
    # 配置按钮框架的列权重以实现居中
    button_frame.columnconfigure(0, weight=1)
    button_frame.columnconfigure(1, weight=0)
    button_frame.columnconfigure(2, weight=0)
    button_frame.columnconfigure(3, weight=1)
    
    def on_yes():
        result.set(True)
        dialog.destroy()
    
    def on_no():
        result.set(False)
        dialog.destroy()
    
    # 居中对称的按钮布局
    ttk.Button(button_frame, text="是", command=on_yes, width=8).grid(row=0, column=1, padx=10)
    ttk.Button(button_frame, text="否", command=on_no, width=8).grid(row=0, column=2, padx=10)
    
    # 根据内容调整窗口大小
    dialog.update_idletasks()
    
    # 计算合适的窗口大小
    req_width = max(500, text_widget.winfo_reqwidth() + 100)
    req_height = max(250, text_widget.winfo_reqheight() + 120)
    
    # 居中显示
    x = (dialog.winfo_screenwidth() // 2) - (req_width // 2)
    y = (dialog.winfo_screenheight() // 2) - (req_height // 2)
    dialog.geometry(f"{req_width}x{req_height}+{x}+{y}")
    
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
        if show_confirm_dialog("确认", "确定要清空所有对话记录吗？"):
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
        self.key_columns = []  # 多关键列列表（新功能）
        self.main_file_name = ""
        self.main_file_path = ""  # 保存主表格文件路径
        self.main_sheet_name = ""  # 保存主表格工作表名称
        self.selected_columns = set()  # 用户选择的参与对比的列
        
        # 工作表选择相关变量
        self.main_sheet_names = []
        self.main_sheet_var = tk.StringVar()
        self.compare_sheet_names = []
        self.compare_sheet_var = tk.StringVar()
        
    def reset(self):
        """重置引擎状态到初始状态"""
        self.main_table = None
        self.main_columns = []
        self.key_column_index = 0
        self.key_columns = []  # 重置多关键列列表
        self.main_file_name = ""
        self.main_file_path = ""  # 重置主表格文件路径
        self.main_sheet_name = ""  # 重置主表格工作表名称
        self.selected_columns = set()
        
    def load_main_table(self, file_path):
        """加载主对比表格"""
        try:
            self.main_file_path = file_path  # 保存文件路径
            self.main_file_name = os.path.splitext(os.path.basename(file_path))[0]
            
            # 读取Excel文件（支持多sheet，这里取第一个）
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            
            # 使用第一个sheet作为主表格
            first_sheet = sheet_names[0]
            self.main_sheet_name = first_sheet  # 保存工作表名称
            
            # 智能数据类型处理：对标识符列强制字符串类型
            dtype_dict = {}
            # 先读取列名
            temp_df = pd.read_excel(file_path, sheet_name=first_sheet, nrows=0)
            for col in temp_df.columns:
                if self._is_identifier_column(col):
                    dtype_dict[col] = str
            
            # 使用智能类型映射读取数据
            if dtype_dict:
                self.main_table = pd.read_excel(file_path, sheet_name=first_sheet, dtype=dtype_dict)
            else:
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
    
    def load_main_file(self, file_path, sheet_name):
        """加载主对比表格的指定工作表"""
        try:
            self.main_file_path = file_path  # 保存文件路径
            self.main_file_name = os.path.splitext(os.path.basename(file_path))[0]
            self.main_sheet_name = sheet_name  # 保存工作表名称
            
            # 智能数据类型处理：对标识符列强制字符串类型
            dtype_dict = {}
            # 先读取列名
            temp_df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=0)
            for col in temp_df.columns:
                if self._is_identifier_column(col):
                    dtype_dict[col] = str
            
            # 使用智能类型映射读取数据
            if dtype_dict:
                self.main_table = pd.read_excel(file_path, sheet_name=sheet_name, dtype=dtype_dict)
            else:
                self.main_table = pd.read_excel(file_path, sheet_name=sheet_name)
            
            self.main_columns = list(self.main_table.columns)
            
            message = f"主表格工作表加载成功: {os.path.basename(file_path)} - {sheet_name}\n"
            message += f"行数: {len(self.main_table)}, 列数: {len(self.main_columns)}\n"
            message += f"列名: {', '.join(self.main_columns[:5])}{'...' if len(self.main_columns) > 5 else ''}"
            
            return True, message
        except Exception as e:
            return False, f"主表格工作表加载失败: {str(e)}"
    
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
    
    def _is_identifier_column(self, column_name):
        """检测是否为标识符列（编码、版本号等）"""
        if not column_name:
            return False
        
        column_name_lower = str(column_name).lower()
        
        # 中文关键词
        identifier_keywords_cn = [
            '编码', '编号', '版本', '型号', '规格', '代码', '代号',
            '序号', '流水号', '工号', '学号', '会员号', '订单号',
            '产品编码', '商品编码', '物料编码', '零件编码'
        ]
        
        # 英文关键词
        identifier_keywords_en = [
            'id', 'code', 'number', 'version', 'model', 'serial',
            'identifier', 'reference', 'sku', 'barcode'
        ]
        
        # 检查中文关键词
        for keyword in identifier_keywords_cn:
            if keyword in column_name:
                return True
        
        # 检查英文关键词
        for keyword in identifier_keywords_en:
            if keyword in column_name_lower:
                return True
        
        return False
    
    def _extract_chinese_chars(self, text):
        """提取文本中的中文字符"""
        import re
        chinese_chars = re.findall(r'[\u4e00-\u9fff]', str(text))
        return ''.join(chinese_chars)
    
    def _get_last_four_chinese(self, text):
        """获取文本中的后四个中文字符"""
        chinese_chars = self._extract_chinese_chars(text)
        return chinese_chars[-4:] if len(chinese_chars) >= 4 else chinese_chars
    
    def _columns_match_by_suffix(self, col1, col2):
        """检查两个列名是否通过后四个中文字符匹配"""
        suffix1 = self._get_last_four_chinese(col1)
        suffix2 = self._get_last_four_chinese(col2)
        
        # 只有当两个列名都至少有4个中文字符且后四个字符相同时才匹配
        return (len(suffix1) >= 4 and len(suffix2) >= 4 and 
                suffix1 == suffix2 and suffix1 != "")
    
    def _columns_match_by_contains(self, col1, col2):
        """检查两个列名是否通过包含关系匹配"""
        str1 = str(col1).strip()
        str2 = str(col2).strip()
        
        # 空字符串不匹配
        if not str1 or not str2:
            return False
        
        # 完全相同不在这里处理
        if str1 == str2:
            return False
        
        # 检查包含关系（较短的是较长的子字符串）
        if len(str1) > len(str2):
            longer, shorter = str1, str2
        else:
            longer, shorter = str2, str1
        
        # 较短的字符串长度至少为2个字符才进行包含匹配
        if len(shorter) < 2:
            return False
        
        return shorter in longer
    
    def _extract_keywords(self, text):
        """提取列名中的关键词"""
        import re
        text = str(text).strip()
        
        # 定义常见的关键词模式
        keywords = set()
        
        # 提取中文关键词（2-6个字符的中文词组）
        chinese_words = re.findall(r'[\u4e00-\u9fff]{2,6}', text)
        keywords.update(chinese_words)
        
        # 提取英文关键词（2个字符以上的英文单词，但排除常见单词）
        english_words = re.findall(r'[a-zA-Z]{2,}', text)
        # 排除常见的无意义单词
        excluded_words = {'id', 'no', 'of', 'in', 'on', 'at', 'to', 'for', 'the', 'and', 'or'}
        english_words = [word.lower() for word in english_words if word.lower() not in excluded_words]
        keywords.update(english_words)
        
        # 提取数字关键词
        numbers = re.findall(r'\d+', text)
        keywords.update(numbers)
        
        # 排除过于常见的中文词汇（减少误匹配）
        excluded_chinese = {'的列', '列', '的', '个', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'}
        keywords = keywords - excluded_chinese
        
        return keywords
    
    def _columns_match_by_keywords(self, col1, col2):
        """检查两个列名是否通过关键词匹配"""
        keywords1 = self._extract_keywords(col1)
        keywords2 = self._extract_keywords(col2)
        
        # 需要至少有一个关键词，且关键词不能太短
        if not keywords1 or not keywords2:
            return False
        
        # 计算关键词交集
        common_keywords = keywords1 & keywords2
        
        # 过滤有效的共同关键词
        valid_common = set()
        for kw in common_keywords:
            kw_str = str(kw)
            # 中文关键词至少2个字符，英文关键词至少3个字符
            if (len(kw_str) >= 2 and any('\u4e00' <= c <= '\u9fff' for c in kw_str)) or len(kw_str) >= 3:
                valid_common.add(kw)
        
        # 至少要有一个有效的共同关键词，且这个关键词要有实际意义
        if not valid_common:
            return False
        
        # 进一步验证：确保关键词确实有意义
        meaningful_keywords = set()
        for kw in valid_common:
            kw_str = str(kw)
            # 名称、编码、金额、数量、价格等有意义的词汇
            if any(meaningful in kw_str for meaningful in ['名称', '编码', '编号', '金额', '数量', '价格', '客户', '产品', '商品', '订单', '仓库']):
                meaningful_keywords.add(kw)
            elif kw_str.lower() in ['id', 'name', 'code', 'amount', 'price', 'customer', 'product', 'order']:
                meaningful_keywords.add(kw)
        
        return len(meaningful_keywords) > 0
    
    def _find_matching_columns(self, main_cols, compare_cols):
        """找到匹配的列，只使用精确匹配"""
        # 只使用精确匹配，完全移除智能匹配逻辑
        return list(set(main_cols) & set(compare_cols))
    
    def get_common_columns_from_files(self, compare_files, selected_sheet=None):
        """获取主表格和所有对比文件的共同列，只使用精确匹配"""
        if self.main_table is None:
            return []
        
        # 只使用精确匹配模式，完全移除智能匹配逻辑
        common_columns = set(self.main_columns)
        
        try:
            for file_idx, file_path in enumerate(compare_files, 1):
                # 添加文件名调试输出
                file_name = os.path.basename(file_path)
                print(f"DEBUG: 文件{file_idx} 文件名: {file_name}")
                
                try:
                    # 只使用第一个工作表进行比较，与实际对比逻辑保持一致
                    excel_file = pd.ExcelFile(file_path)
                    sheet_names = excel_file.sheet_names
                    
                    if sheet_names:
                        # 使用选定的工作表，如果指定且存在的话
                        if selected_sheet and selected_sheet.strip() and selected_sheet in sheet_names:
                            first_sheet = selected_sheet
                        else:
                            first_sheet = sheet_names[0]
                        
                        # 使用智能表头检测
                        header_row, columns = self._find_header_row(file_path, first_sheet)
                        
                        if header_row is not None:
                            # 智能数据类型处理
                            dtype_dict = {}
                            temp_df = pd.read_excel(file_path, sheet_name=first_sheet, header=header_row, nrows=0)
                            for col in temp_df.columns:
                                if self._is_identifier_column(col):
                                    dtype_dict[col] = str
                            
                            # 使用检测到的表头行读取数据
                            if dtype_dict:
                                compare_df = pd.read_excel(file_path, sheet_name=first_sheet, header=header_row, dtype=dtype_dict)
                            else:
                                compare_df = pd.read_excel(file_path, sheet_name=first_sheet, header=header_row)
                            compare_columns = set(compare_df.columns)
                            print(f"DEBUG: 文件{file_idx} 使用智能表头检测，表头在第{header_row + 1}行")
                        else:
                            # 如果智能检测失败，使用默认方式读取
                            print(f"DEBUG: 文件{file_idx} 智能检测失败，使用默认方式读取")
                            dtype_dict = {}
                            temp_df = pd.read_excel(file_path, sheet_name=first_sheet, nrows=0)
                            for col in temp_df.columns:
                                if self._is_identifier_column(col):
                                    dtype_dict[col] = str
                            
                            if dtype_dict:
                                compare_df = pd.read_excel(file_path, sheet_name=first_sheet, dtype=dtype_dict)
                            else:
                                compare_df = pd.read_excel(file_path, sheet_name=first_sheet)
                            compare_columns = set(compare_df.columns)
                        
                        # Bug #3修复：检查文件是否有有效列名，跳过无效文件而不是提前退出
                        valid_columns = set()
                        for col in compare_columns:
                            col_str = str(col).strip()
                            # 排除空列名、Unnamed列、纯数字列名等无效列名
                            if (col_str and
                                not col_str.startswith('Unnamed:') and
                                not col_str.isdigit() and
                                len(col_str) > 1):
                                valid_columns.add(col)
                        
                        # 显示有效列名（取前5个用于调试）
                        valid_cols_list = sorted(list(valid_columns))
                        cols_preview = valid_cols_list[:5] if len(valid_cols_list) > 5 else valid_cols_list
                        print(f"DEBUG: 文件{file_idx} 有效列名: {cols_preview}...")
                        print(f"DEBUG: 文件{file_idx} 交集前: {len(common_columns)}, 交集后: {len(common_columns & valid_columns)}")
                        
                        # 如果当前文件有有效列，才参与交集计算
                        if valid_columns:
                            # 与当前文件的有效列取交集
                            common_columns = common_columns & valid_columns
                        else:
                            print(f"DEBUG: 文件{file_idx} 无有效列名，跳过此文件")
                            continue
                        
                except Exception as e:
                    # 如果读取文件失败，跳过这个文件
                    continue
                    
            return sorted(list(common_columns))
        except Exception as e:
            return []
    
    def _normalize_value(self, value):
        """标准化值，处理数据类型不一致但值相同的情况，保持前导零"""
        if pd.isna(value):
            return None
        
        # 转换为字符串并去掉首尾空格
        str_value = str(value).strip()
        
        # 处理字符串形式的 'nan'
        if str_value.lower() in ['nan', 'none', '']:
            return None
        
        # 对于包含前导零的字符串，直接返回字符串形式以保持前导零
        # 检查是否是以0开头的数字字符串（但不是单独的"0"）
        if str_value.startswith('0') and len(str_value) > 1 and str_value.isdigit():
            return str_value  # 保持前导零
        
        # 对于其他情况，尝试转换为数字（处理数值类型不一致的情况）
        try:
            # 如果是整数形式的浮点数，转换为整数
            if '.' in str_value and str_value.replace('.', '').isdigit():
                float_val = float(str_value)
                if float_val.is_integer():
                    return int(float_val)
                return float_val
            # 如果是纯数字（但不以0开头），转换为整数
            elif str_value.isdigit():
                return int(str_value)
            # 尝试转换为浮点数
            elif str_value.replace('.', '').replace('-', '').isdigit():
                return float(str_value)
        except:
            pass
        
        return str_value
    
    def _normalize_value_by_reference(self, value, reference_value):
        """根据参考值（主表格值）的格式来标准化对比值"""
        if pd.isna(value):
            return None
        if pd.isna(reference_value):
            return self._normalize_value(value)
        
        # 转换为字符串并去掉首尾空格
        str_value = str(value).strip()
        str_ref = str(reference_value).strip()
        
        # 处理字符串形式的 'nan'
        if str_value.lower() in ['nan', 'none', '']:
            return None
        
        # 对于包含前导零的字符串，直接返回字符串形式以保持前导零
        if str_value.startswith('0') and len(str_value) > 1 and str_value.isdigit():
            return str_value  # 保持前导零
        
        # 如果参考值是数值类型，尝试按照参考值的格式来处理
        try:
            # 检查参考值的类型和格式
            if isinstance(reference_value, (int, float)):
                # 参考值是数值类型
                if isinstance(reference_value, int):
                    # 参考值是整数，尝试将对比值转为整数
                    if '.' in str_value:
                        float_val = float(str_value)
                        if float_val.is_integer():
                            return int(float_val)
                        # 如果不是整数，保持浮点数但尽量简化
                        return float_val
                    else:
                        return int(str_value) if str_value.isdigit() else str_value
                elif isinstance(reference_value, float):
                    # 参考值是浮点数，检查其小数位数
                    ref_decimal_places = 0
                    if '.' in str_ref:
                        ref_decimal_part = str_ref.split('.')[1]
                        # 去掉尾部的0
                        ref_decimal_part = ref_decimal_part.rstrip('0')
                        ref_decimal_places = len(ref_decimal_part)
                    
                    # 将对比值转为浮点数并按照参考值的小数位数格式化
                    float_val = float(str_value)
                    if ref_decimal_places == 0:
                        # 参考值没有小数部分，检查对比值是否为整数
                        if float_val.is_integer():
                            return int(float_val)
                        else:
                            # 保持原始精度，但限制显示位数
                            return round(float_val, 6)  # 最多6位小数
                    else:
                        # 按照参考值的小数位数来格式化
                        return round(float_val, ref_decimal_places)
            else:
                # 参考值是字符串类型，检查是否为数值字符串
                if '.' in str_ref and str_ref.replace('.', '').replace('-', '').isdigit():
                    # 参考值是浮点数字符串
                    ref_decimal_places = len(str_ref.split('.')[1].rstrip('0'))
                    float_val = float(str_value)
                    if ref_decimal_places == 0:
                        return int(float_val) if float_val.is_integer() else float_val
                    else:
                        return round(float_val, ref_decimal_places)
                elif str_ref.isdigit():
                    # 参考值是整数字符串
                    float_val = float(str_value)
                    return int(float_val) if float_val.is_integer() else float_val
        except:
            pass
        
        # 如果无法按照参考值格式化，使用原始标准化方法
        return self._normalize_value(value)
    
    def _values_are_different(self, val1, val2, use_reference_format=False):
        """判断两个值是否不同，处理各种边界情况"""
        if use_reference_format:
            # 使用val1（主表格值）作为参考来标准化val2（对比表格值）
            norm_val1 = self._normalize_value(val1)
            norm_val2 = self._normalize_value_by_reference(val2, val1)
        else:
            # 传统方式：分别标准化
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
    
    def compare_tables(self, compare_files, output_dir, progress_callback=None, selected_sheet=None):
        """执行表格对比 - 改进版本，支持多关键列"""
        if self.main_table is None:
            return False, "请先加载主表格"
        
        if not compare_files:
            return False, "请选择待对比文件"
        
        # 获取有效的关键列
        effective_key_columns = self._get_effective_key_columns()
        if not effective_key_columns:
            return False, "未选择关键列，请至少选择一个关键列"
        
        # 验证关键列是否存在于主表格中
        missing_keys = [col for col in effective_key_columns if col not in self.main_columns]
        if missing_keys:
            return False, f"关键列 {missing_keys} 在主表格中不存在"
        
        # 显示关键列信息
        if len(effective_key_columns) == 1:
            key_info = f"单关键列: '{effective_key_columns[0]}'"
        else:
            key_info = f"多关键列联合匹配: {effective_key_columns}"
        
        if progress_callback:
            progress_callback(f"开始对比，{key_info}")
        
        all_difference_records = []
        total_files = len(compare_files)
        total_differences = 0
        
        for file_idx, compare_file_path in enumerate(compare_files, 1):
            compare_file_name = os.path.splitext(os.path.basename(compare_file_path))[0]
            
            # 添加文件名调试输出
            print(f"DEBUG: 对比文件{file_idx} 文件名: {os.path.basename(compare_file_path)}")
            
            if progress_callback:
                progress_callback(f"[{file_idx}/{total_files}] 处理文件: {os.path.basename(compare_file_path)}")
            
            try:
                # 读取对比文件
                excel_file = pd.ExcelFile(compare_file_path)
                sheet_names = excel_file.sheet_names
                
                # 确定要处理的工作表
                if selected_sheet and selected_sheet in sheet_names:
                    # 使用用户选择的工作表
                    target_sheets = [selected_sheet]
                else:
                    # 如果没有指定或指定的工作表不存在，使用第一个工作表
                    target_sheets = [sheet_names[0]] if sheet_names else []
                
                for sheet_name in target_sheets:
                    if progress_callback:
                        progress_callback(f"  处理工作表: {sheet_name}")
                    
                    try:
                        # 使用智能表头检测
                        header_row, columns = self._find_header_row(compare_file_path, sheet_name)
                        
                        if header_row is not None:
                            # 智能数据类型处理
                            dtype_dict = {}
                            temp_df = pd.read_excel(compare_file_path, sheet_name=sheet_name, header=header_row, nrows=0)
                            for col in temp_df.columns:
                                if self._is_identifier_column(col):
                                    dtype_dict[col] = str
                            
                            # 使用检测到的表头行读取数据
                            if dtype_dict:
                                compare_df = pd.read_excel(compare_file_path, sheet_name=sheet_name, header=header_row, dtype=dtype_dict)
                            else:
                                compare_df = pd.read_excel(compare_file_path, sheet_name=sheet_name, header=header_row)
                        else:
                            # 如果智能检测失败，使用默认方式读取
                            print(f"DEBUG: 对比文件{file_idx} 工作表{sheet_name} 使用默认方式读取")
                            dtype_dict = {}
                            temp_df = pd.read_excel(compare_file_path, sheet_name=sheet_name, nrows=0)
                            for col in temp_df.columns:
                                if self._is_identifier_column(col):
                                    dtype_dict[col] = str
                            
                            if dtype_dict:
                                compare_df = pd.read_excel(compare_file_path, sheet_name=sheet_name, dtype=dtype_dict)
                            else:
                                compare_df = pd.read_excel(compare_file_path, sheet_name=sheet_name)
                        
                        # 应用智能列匹配：标准化列顺序
                        # ⚠️ 完全移除智能匹配逻辑：不再调整列顺序，严格按列名称比对
                        if progress_callback:
                            progress_callback(f"    📋 保持原始列顺序，按列名称进行比对...")
                        
                        # 检查关键列是否存在
                        missing_key_columns = [col for col in effective_key_columns if col not in compare_df.columns]
                        if missing_key_columns:
                            if progress_callback:
                                progress_callback(f"    ⚠ 跳过：关键列 {missing_key_columns} 不存在")
                            continue
                        
                        # 找到两个表格的共同列（支持智能匹配），排除临时列
                        main_columns_filtered = [col for col in self.main_table.columns if col != '__original_row_number__']
                        compare_columns_filtered = [col for col in compare_df.columns if col != '__original_row_number__']
                        
                        # ⚠️ 完全重写：严格按照列名称进行比对，不使用智能匹配
                        # 用户要求：只对列名称完全相同的列进行比对，不管位置如何
                        main_columns_filtered = [col for col in self.main_table.columns if col != '__original_row_number__']
                        compare_columns_filtered = [col for col in compare_df.columns if col != '__original_row_number__']
                        
                        # 只对完全同名的列进行内容比对（精确匹配，不使用智能匹配）
                        common_columns = list(set(main_columns_filtered) & set(compare_columns_filtered))
                        
                        if progress_callback:
                            progress_callback(f"    📋 主表格列: {main_columns_filtered[:5]}{'...' if len(main_columns_filtered) > 5 else ''}")
                            progress_callback(f"    📋 对比表格列: {compare_columns_filtered[:5]}{'...' if len(compare_columns_filtered) > 5 else ''}")
                            progress_callback(f"    ✅ 完全同名列 ({len(common_columns)} 个): {common_columns}")
                        
                        # 使用用户选择的列（从同名列中筛选）
                        if self.selected_columns:
                            # 只使用用户选择的同名列
                            original_common_count = len(common_columns)
                            common_columns = [col for col in common_columns if col in self.selected_columns]
                            
                            if progress_callback:
                                progress_callback(f"    🎯 用户选择了 {len(self.selected_columns)} 列，其中 {len(common_columns)} 列为同名列")
                                progress_callback(f"    ✅ 最终比对列: {common_columns}")
                        
                        if not common_columns:
                            if progress_callback:
                                progress_callback(f"    ⚠ 跳过：没有可比对的同名列")
                            continue
                        
                        
                        # 在设置索引之前，先保存原始行号信息
                        main_table_with_row_num = self.main_table.copy()
                        main_table_with_row_num['__original_row_number__'] = range(2, len(self.main_table) + 2)  # Excel行号从2开始（考虑表头）
                        
                        compare_df_with_row_num = compare_df.copy()
                        compare_df_with_row_num['__original_row_number__'] = range(2, len(compare_df) + 2)  # Excel行号从2开始（考虑表头）
                        
                        # 为多关键列创建联合键
                        if len(effective_key_columns) > 1:
                            # 多关键列：生成联合键
                            main_table_with_row_num['__multi_key__'] = main_table_with_row_num.apply(
                                lambda row: self._create_multi_key(row, effective_key_columns), axis=1)
                            compare_df_with_row_num['__multi_key__'] = compare_df_with_row_num.apply(
                                lambda row: self._create_multi_key(row, effective_key_columns), axis=1)
                            
                            # 使用联合键作为索引
                            main_indexed = main_table_with_row_num.set_index('__multi_key__')
                            compare_indexed = compare_df_with_row_num.set_index('__multi_key__')
                            key_column_name = "联合键"  # 用于显示
                        else:
                            # 单关键列：直接使用该列作为索引  
                            key_column_name = effective_key_columns[0]
                            main_indexed = main_table_with_row_num.set_index(key_column_name)
                            compare_indexed = compare_df_with_row_num.set_index(key_column_name)
                        
                        # 找到需要对比的行（关键列值相同的行）
                        common_keys = set(main_indexed.index) & set(compare_indexed.index)
                        
                        if not common_keys:
                            if progress_callback:
                                progress_callback(f"    ⚠ 跳过：没有匹配的关键列值")
                            continue
                        
                        sheet_differences = 0
                        
                        # 逐行对比
                        for key in common_keys:
                            # 处理重复键的情况：如果键重复，只取第一个
                            # 注意：当键唯一时，loc[key]返回Series；当键重复时，返回DataFrame
                            if isinstance(main_indexed.loc[key], pd.DataFrame):
                                # 如果有重复键，取第一个
                                main_row = main_indexed.loc[key].iloc[0]
                                main_row_number = main_row['__original_row_number__']
                            else:
                                # 键是唯一的，直接使用
                                main_row = main_indexed.loc[key]
                                main_row_number = main_row['__original_row_number__']
                            
                            if isinstance(compare_indexed.loc[key], pd.DataFrame):
                                # 如果有重复键，取第一个
                                compare_row = compare_indexed.loc[key].iloc[0]
                                compare_row_number = compare_row['__original_row_number__']
                            else:
                                # 键是唯一的，直接使用
                                compare_row = compare_indexed.loc[key]
                                compare_row_number = compare_row['__original_row_number__']
                            
                            # 检查共同列是否有差异（使用列名映射）
                            row_has_differences = False
                            different_columns = []
                            
                            # 核心比对逻辑：严格按照列名称进行比对
                            # 1. 从主比对表格选择一个关键列中的一行内容 ✓ (当前行：main_row)
                            # 2. 从待比对表格中筛选出相同关键列的行内容 ✓ (当前行：compare_row)
                            # 3. 把关键列除外的相同列名称的行内容进行比对
                            
                            for column_name in common_columns:
                                if column_name in effective_key_columns:
                                    continue  # 跳过关键列本身
                                
                                # 严格按照列名称获取值，不依赖位置
                                main_value = None
                                compare_value = None
                                
                                # 从主表格获取该列名的值
                                if column_name in main_row.index:
                                    main_value = main_row[column_name]
                                
                                # 从对比表格获取同名列的值
                                if column_name in compare_row.index:
                                    compare_value = compare_row[column_name]
                                
                                # 比对值是否不同（使用主表格值作为参考格式）
                                if self._values_are_different(main_value, compare_value, use_reference_format=True):
                                    row_has_differences = True
                                    different_columns.append(column_name)
                            
                            # 如果有差异，记录到结果中
                            if row_has_differences:
                                # 创建主表格记录（上一行）
                                main_record = {}
                                
                                # 处理关键列显示
                                if len(effective_key_columns) == 1:
                                    main_record[effective_key_columns[0]] = key
                                else:
                                    # 多关键列：分别显示每个关键列的值
                                    for i, col in enumerate(effective_key_columns):
                                        if col in main_row.index:
                                            main_record[col] = main_row[col]
                                        else:
                                            main_record[col] = ''
                                
                                for main_col in common_columns:
                                    if main_col not in effective_key_columns:
                                        main_value = main_row[main_col] if main_col in main_row.index else ''
                                        # 对主表格值进行标准化
                                        normalized_main_value = self._normalize_value(main_value)
                                        main_record[main_col] = normalized_main_value if normalized_main_value is not None else main_value
                                main_record['差异列'] = ', '.join(different_columns)
                                # 使用之前已经计算好的行号（在第275-286行计算）
                                main_record['数据来源'] = f'{self.main_file_name}_{self.main_sheet_name}_主表格_行{main_row_number}'
                                all_difference_records.append(main_record)
                                
                                # 创建对比表格记录（下一行）- 使用主表格列名作为标准
                                compare_record = {}
                                
                                # 处理关键列显示
                                if len(effective_key_columns) == 1:
                                    compare_record[effective_key_columns[0]] = key
                                else:
                                    # 多关键列：分别显示每个关键列的值
                                    for i, col in enumerate(effective_key_columns):
                                        if col in compare_row.index:
                                            compare_record[col] = compare_row[col]
                                        else:
                                            compare_record[col] = ''
                                
                                # 严格按照列名称填写数据列内容，按主表格格式标准化
                                # 添加调试信息来检测列内容错位问题
                                debug_info = []
                                for main_col in common_columns:
                                    if main_col not in effective_key_columns:
                                        # 直接按照列名称获取值，不使用column_mapping
                                        if main_col in compare_row.index:
                                            compare_value = compare_row[main_col]
                                            # 获取主表格中对应列的值作为参考
                                            main_reference_value = main_row[main_col] if main_col in main_row.index else None
                                            
                                            # 调试信息：检测可能的列内容错位
                                            debug_info.append(f"{main_col}: 主表格='{main_reference_value}', 对比表格='{compare_value}'")
                                            
                                            # 按照主表格的格式标准化对比值
                                            normalized_compare_value = self._normalize_value_by_reference(compare_value, main_reference_value)
                                            compare_record[main_col] = normalized_compare_value if normalized_compare_value is not None else compare_value
                                        else:
                                            compare_record[main_col] = ''
                                
                                # 输出调试信息（仅在发现差异时）
                                if debug_info:
                                    print(f"DEBUG: 关键值 {key} 的列内容对比:")
                                    for info in debug_info[:3]:  # 只显示前3个列的信息
                                        print(f"  {info}")
                                
                                compare_record['差异列'] = ', '.join(different_columns)
                                # 使用之前已经计算好的行号（在第292-303行计算）
                                compare_record['数据来源'] = f'{compare_file_name}_{sheet_name}_对比表格_行{compare_row_number}'
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
            
            # 获取有效的关键列
            effective_key_columns = self._get_effective_key_columns()
            if not effective_key_columns:
                return False, "错误：没有设置关键列", None
            
            # 确定用于分组的关键列（使用第一个关键列进行分组）
            primary_key_column = effective_key_columns[0]
            
            # 按关键列值分组
            main_records = {}
            compare_records = {}
            
            for record in difference_records:
                # 如果是多关键列，使用联合键；如果是单关键列，使用该列的值
                if len(effective_key_columns) > 1:
                    # 多关键列：生成联合键用于分组
                    key_values = []
                    for col in effective_key_columns:
                        if col in record:
                            key_values.append(str(record[col]))
                        else:
                            key_values.append("")
                    key_value = "|||".join(key_values)
                else:
                    # 单关键列：直接使用该列的值
                    key_value = record.get(primary_key_column, "")
                
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
            if primary_key_column in data_columns:
                data_columns.remove(primary_key_column)
            final_columns = [primary_key_column] + sorted(data_columns) + system_columns
            
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
- 关键列: {effective_key_columns if len(effective_key_columns) > 1 else effective_key_columns[0]}
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

    def _find_header_row(self, file_path, sheet_name, max_check_rows=10):
        """
        智能检测Excel文件的表头行位置
        返回: (header_row_index, columns_list) 或 (None, None)
        """
        try:
            # 读取前几行数据来检测表头
            for header_row in range(max_check_rows):
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, dtype=str)
                    columns = list(df.columns)
                    
                    # 检查列名的有效性
                    valid_columns = []
                    invalid_count = 0
                    
                    for col in columns:
                        col_str = str(col).strip()
                        # 更严格的有效性检查
                        if (col_str and 
                            col_str != 'nan' and
                            not col_str.startswith('Unnamed') and  # 更严格，不用:
                            not col_str.isdigit() and 
                            not col_str.replace('.', '').isdigit() and  # 排除小数
                            len(col_str) > 1 and
                            col_str not in ['nan', 'None', 'NULL', '空', '']):  # 排除更多无效值
                            valid_columns.append(col_str)
                        else:
                            invalid_count += 1
                    
                    # 如果有效列数量占总列数的60%以上，认为找到了表头行
                    total_cols = len(columns)
                    valid_ratio = len(valid_columns) / total_cols if total_cols > 0 else 0
                    
                    if valid_ratio >= 0.6 and len(valid_columns) >= 3:  # 至少3个有效列且占比60%以上
                        print(f"DEBUG: 文件 {os.path.basename(file_path)} 工作表 {sheet_name} 找到表头在第{header_row + 1}行")
                        print(f"DEBUG: 有效列名: {valid_columns[:5]}...")  # 显示前5个有效列名
                        return header_row, valid_columns
                        
                except Exception as e:
                    continue
            
            # 如果没找到合适的表头行，返回None
            print(f"DEBUG: 文件 {os.path.basename(file_path)} 工作表 {sheet_name} 未找到有效表头")
            return None, None
            
        except Exception as e:
            print(f"DEBUG: 文件 {os.path.basename(file_path)} 表头检测出错: {str(e)}")
            return None, None

    def _create_multi_key(self, row, key_columns):
        """
        生成多关键列的联合键
        参数:
            row: pandas Series，数据行
            key_columns: list，关键列名列表
        返回:
            str: 联合键字符串
        """
        if not key_columns:
            return None
        
        # 将多个关键列的值用特殊分隔符连接
        key_values = []
        for col in key_columns:
            if col in row.index:
                value = str(row[col]).strip() if pd.notna(row[col]) else ""
                key_values.append(value)
            else:
                key_values.append("")  # 如果列不存在，使用空字符串
        
        return "|||".join(key_values)  # 使用特殊分隔符避免冲突
    
    def _get_effective_key_columns(self):
        """
        获取有效的关键列
        返回:
            list: 关键列列表（多列或单列）
        """
        if self.key_columns:
            # 如果有多关键列设置，优先使用
            return self.key_columns
        elif hasattr(self, 'key_column_index') and self.main_columns:
            # 向后兼容：如果没有多关键列但有单关键列索引，转换为列表
            if 0 <= self.key_column_index < len(self.main_columns):
                return [self.main_columns[self.key_column_index]]
        return []

class ImprovedExcelCompareGUI:
    """改进版Excel对比工具GUI界面"""
    
    def __init__(self, root):
        self.root = root
        self.engine = ImprovedExcelCompareEngine()
        self.compare_files = []
        
        # 工作表选择相关变量
        self.main_sheet_names = []
        self.main_sheet_var = tk.StringVar()
        self.compare_sheet_names = []
        self.compare_sheet_var = tk.StringVar()
        
        # 多关键列相关变量
        self.key_columns = []  # 存储选中的关键列列表
        
        self.setup_ui()
        
    def setup_ui(self):
        """设置用户界面"""
        self.root.title("Excel小工具 v4.3")
        self.root.geometry("1200x900")  # 增加窗口初始大小
        
        # 设置最小窗口大小，确保界面不会过小
        self.root.minsize(1100, 800)  # 增加最小窗口大小，确保所有控件都能显示
        
        # 设置主题
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except:
            pass
        
        # 设置绿色进度条样式
        style.configure("green.Horizontal.TProgressbar", 
                       background='green', 
                       troughcolor='lightgray',
                       borderwidth=1)
        
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
        title_label = ttk.Label(main_frame, text="Excel小工具 v4.3",
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
        desc_text = "特点: • 指定列精确行匹配 | • 智能选列(关键列必选) | • 类型差异智能处理 | • 竖排差异红粗高亮 | • 弹窗选列 | • 自动序号报告 | • 并排优化 | • 精简流程"
        desc_label = ttk.Label(compare_frame, text=desc_text, font=("Arial", 9), 
                              foreground="darkblue", anchor="w", justify="left") # 添加 anchor 和 justify
        desc_label.grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="ew") # 修改 pady 和 sticky
        
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
        
        # 主表格工作表选择
        ttk.Label(main_file_frame, text="工作表:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.main_sheet_combo = ttk.Combobox(main_file_frame, textvariable=self.main_sheet_var,
                                           state="readonly", font=("Arial", 9))
        self.main_sheet_combo.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=(0, 10), pady=(0, 10))
        self.main_sheet_combo.bind('<<ComboboxSelected>>', self.on_main_sheet_selected)
        
        # 右侧：关键列选择
        key_column_frame = ttk.LabelFrame(compare_frame, text="2. 选择用于行匹配的关键列", padding="10")
        key_column_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        key_column_frame.columnconfigure(0, weight=1)
        key_column_frame.rowconfigure(1, weight=1)
        
        # 关键列选择说明
        ttk.Label(key_column_frame, text="支持单列或多列联合匹配:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # 关键列列表框架
        key_list_frame = ttk.Frame(key_column_frame)
        key_list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        key_list_frame.columnconfigure(0, weight=1)
        key_list_frame.rowconfigure(0, weight=1)
        
        # 已选关键列列表
        self.key_columns_tree = ttk.Treeview(key_list_frame, columns=('column',), show='headings', height=4)
        self.key_columns_tree.heading('column', text='已选关键列')
        self.key_columns_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 关键列滚动条
        key_scrollbar = ttk.Scrollbar(key_list_frame, orient=tk.VERTICAL, command=self.key_columns_tree.yview)
        key_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.key_columns_tree.configure(yscrollcommand=key_scrollbar.set)
        
        # 关键列操作按钮
        key_buttons_frame = ttk.Frame(key_column_frame)
        key_buttons_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        key_buttons_frame.columnconfigure(1, weight=1)
        
        # 列选择下拉框
        ttk.Label(key_buttons_frame, text="添加列:").grid(row=0, column=0, padx=(0, 5))
        self.key_column_var = tk.StringVar()
        self.key_column_combo = ttk.Combobox(key_buttons_frame, textvariable=self.key_column_var,
                                           state="readonly", font=("Arial", 9))
        self.key_column_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # 操作按钮
        key_op_frame = ttk.Frame(key_buttons_frame)
        key_op_frame.grid(row=0, column=2)
        
        ttk.Button(key_op_frame, text="添加", 
                  command=self.add_key_column).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(key_op_frame, text="移除", 
                  command=self.remove_key_column).grid(row=0, column=1, padx=(0, 5))
        ttk.Button(key_op_frame, text="清空", 
                  command=self.clear_key_columns).grid(row=0, column=2)
        
        # 显示关键列信息
        self.key_column_info = tk.StringVar(value="请先选择主表格")
        ttk.Label(key_column_frame, textvariable=self.key_column_info, 
                 foreground="darkgreen", font=("Arial", 9), wraplength=250).grid(row=3, column=0, 
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
                  command=self.clear_files).grid(row=2, column=0, pady=(0, 5))
        
        # 对比文件工作表选择
        ttk.Label(files_buttons_frame, text="默认工作表:", font=("Arial", 8)).grid(row=3, column=0, sticky=tk.W, pady=(10, 2))
        self.compare_sheet_combo = ttk.Combobox(files_buttons_frame, textvariable=self.compare_sheet_var,
                                              state="readonly", font=("Arial", 8), width=12)
        self.compare_sheet_combo.grid(row=4, column=0, pady=(0, 5))
        self.compare_sheet_combo.bind('<<ComboboxSelected>>', self.on_compare_sheet_selected)
        
        # 右侧：比对设置（竖排布局）
        settings_frame = ttk.LabelFrame(compare_frame, text="4. 比对设置", padding="10")
        settings_frame.grid(row=2, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        settings_frame.columnconfigure(0, weight=1)
        # 修复权重配置：确保两个子区域都有足够的空间
        settings_frame.rowconfigure(0, weight=0, minsize=55)  # 进一步减小 minsize
        settings_frame.rowconfigure(1, weight=0, minsize=50)  # 大幅减小 output_frame 行的 minsize
        
        # 上部：比对列选择
        exclude_frame = ttk.LabelFrame(settings_frame, text="比对列选择", padding="8")
        exclude_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        exclude_frame.columnconfigure(0, weight=1) # Column for the instruction label
        exclude_frame.columnconfigure(1, weight=0) # Column for the button (no weight to keep it compact)
        
        # Row 0: Instruction Label and Button
        exclude_frame.rowconfigure(0, weight=0) 
        ttk.Label(exclude_frame, text="选择哪些列参与对比（关键列为必选项）：", 
                 font=("Arial", 9), foreground="darkblue").grid(row=0, column=0, sticky=tk.W, pady=(0, 2), padx=(0,5)) # pady 변경
        
        self.exclude_expand_button = ttk.Button(exclude_frame, text="选择比对列", 
                                               command=self.open_column_selection_window,
                                               width=15)
        self.exclude_expand_button.grid(row=0, column=1, sticky=tk.E, pady=(0, 8), padx=(0,5))

        # Row 1: Status Information
        exclude_frame.rowconfigure(1, weight=0)
        self.exclude_info = tk.StringVar(value="当前无选中列 - 请先选择主表格和待对比文件")
        status_label = ttk.Label(exclude_frame, textvariable=self.exclude_info, 
                               foreground="darkorange", font=("Arial", 9), wraplength=350) # Increased wraplength
        status_label.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(2, 5)) # pady 변경
        
        # 下部：输出目录选择
        output_frame = ttk.LabelFrame(settings_frame, text="输出目录", padding="8")
        output_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        # Configure columns for three widgets in one row
        output_frame.columnconfigure(0, weight=0) # Label
        output_frame.columnconfigure(1, weight=1) # Entry (stretchable)
        output_frame.columnconfigure(2, weight=0) # Button
        # Remove rowconfigure for row 1 as it's no longer needed
        # output_frame.rowconfigure(0, weight=0) # This was for the label in its own row
        # output_frame.rowconfigure(1, weight=1, minsize=50) # This was for the controls in their own row
        
        ttk.Label(output_frame, text="选择差异报告的保存位置：", 
                 font=("Arial", 9), foreground="darkblue").grid(row=0, column=0, sticky=tk.W, pady=(0, 8), padx=(0,5))
        
        # output_controls_frame = ttk.Frame(output_frame) # No longer needed
        # output_controls_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 8)) # No longer needed
        # output_controls_frame.columnconfigure(0, weight=1) # No longer needed
        
        self.output_dir_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_dir_var, 
                 state="readonly", font=("Arial", 10)).grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=(0,8))
        ttk.Button(output_frame, text="选择目录", 
                  command=self.select_output_dir).grid(row=0, column=2, sticky=tk.E, pady=(0,8), padx=(0,5))
        
        # 5. 执行对比
        action_frame = ttk.Frame(compare_frame)
        action_frame.grid(row=3, column=0, columnspan=2, pady=(0, 15), sticky="ew") # Add sticky="ew"
        
        # Configure columns within action_frame for centering
        action_frame.columnconfigure(0, weight=1) # Flexible space on the left
        action_frame.columnconfigure(1, weight=0) # For the Label
        action_frame.columnconfigure(2, weight=0) # For the '开始对比' button
        action_frame.columnconfigure(3, weight=0) # For the '重置' button
        action_frame.columnconfigure(4, weight=1) # Flexible space on the right

        ttk.Label(action_frame, text="5. 执行对比:", 
                 font=("Arial", 11, "bold")).grid(row=0, column=1, sticky=tk.E, padx=(0, 10)) # Place in col 1, add sticky and adjust padx
        
        self.compare_button = ttk.Button(action_frame, text="开始对比", 
                                        command=self.start_compare, 
                                        style="Accent.TButton")
        self.compare_button.grid(row=0, column=2, padx=(0, 10)) # Place in col 2
        
        ttk.Button(action_frame, text="重置", 
                  command=self.reset).grid(row=0, column=3) # Place in col 3
        
        # 6. 日志区域
        ttk.Label(compare_frame, text="6. 处理日志:", 
                 font=("Arial", 11, "bold")).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        log_frame = ttk.Frame(compare_frame)
        log_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, font=("Consolas", 9))
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 进度条替代状态栏
        progress_frame = ttk.Frame(compare_frame)
        progress_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        progress_frame.columnconfigure(1, weight=1)
        
        ttk.Label(progress_frame, text="进度:", font=("Arial", 9)).grid(row=0, column=0, padx=(0, 10))
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                          maximum=100, style="green.Horizontal.TProgressbar")
        self.progress_bar.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.progress_label = ttk.Label(progress_frame, text="就绪", font=("Arial", 9))
        self.progress_label.grid(row=0, column=2)
        
        # 完善自适应配置 - 设置行列权重以确保各区域能够自适应调整
        # 行权重配置：让需要伸缩的区域有合适的权重
        compare_frame.rowconfigure(0, weight=0)  # 功能说明行，固定高度
        compare_frame.rowconfigure(1, weight=0, minsize=80)  # 主表格和关键列选择行，缩小高度至原来的三分之二
        compare_frame.rowconfigure(2, weight=1, minsize=280)  # 文件列表和设置区域，主要伸缩区域，适当缩小最小高度
        compare_frame.rowconfigure(3, weight=0)  # 执行对比按钮行，固定高度
        compare_frame.rowconfigure(4, weight=0)  # 日志标题行，固定高度
        compare_frame.rowconfigure(5, weight=1, minsize=100)  # 日志区域，次要伸缩区域，缩小最小高度
        compare_frame.rowconfigure(6, weight=0, minsize=30)  # 状态栏和进度条，确保有足够高度显示
        
        # 初始化日志
        self.log("Excel小工具 v4.3 已启动")
        self.log("新功能：标签页界面，支持多种Excel处理工具")
        self.log("对比工具：比对列选择，弹出窗口智能选择，关键列必选机制")
        self.log("界面优化：主表格和关键列并排，待对比文件和比对设置并排")
    
    def setup_process_tab(self):
        """设置表格处理标签页"""
        # 创建表格处理标签页框架
        process_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(process_frame, text="表格处理")
        
        # 配置框架的自适应
        process_frame.columnconfigure(0, weight=1)
        process_frame.columnconfigure(1, weight=1)
        
        # 功能说明
        desc_text = """功能特点：
• 智能表头检测，跳过非表格内容     • 处理多Sheet工作簿，支持隐藏Sheet
• 清除筛选状态，获取完整数据       • 自动检测有效数据起始位置
• 数据清洗与标准化处理           • 支持大文件分块处理
• 并行处理多Sheet提升性能        • 智能编码检测与转换"""
        desc_label = ttk.Label(process_frame, text=desc_text, font=("Arial", 9), 
                              foreground="darkblue")
        desc_label.grid(row=0, column=0, columnspan=2, pady=(0, 15))
        
        # 第一行：文件输入（全宽）
        input_frame = ttk.LabelFrame(process_frame, text="1. 选择Excel文件", padding="10")
        input_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 15))
        input_frame.columnconfigure(0, weight=1)
        input_frame.rowconfigure(0, weight=1)
        
        # 文件列表显示区域
        files_list_frame = ttk.Frame(input_frame)
        files_list_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        files_list_frame.columnconfigure(0, weight=1)
        files_list_frame.rowconfigure(0, weight=1)
        
        # 创建Treeview来显示文件列表
        self.process_files_tree = ttk.Treeview(files_list_frame, columns=('file', 'sheets', 'status'), show='headings', height=4)
        self.process_files_tree.heading('file', text='Excel文件')
        self.process_files_tree.heading('sheets', text='工作表数')
        self.process_files_tree.heading('status', text='状态')
        self.process_files_tree.column('file', width=250)
        self.process_files_tree.column('sheets', width=80)
        self.process_files_tree.column('status', width=100)
        self.process_files_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 滚动条
        process_files_scrollbar = ttk.Scrollbar(files_list_frame, orient=tk.VERTICAL, command=self.process_files_tree.yview)
        process_files_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.process_files_tree.configure(yscrollcommand=process_files_scrollbar.set)
        
        # 文件操作按钮
        files_buttons_frame = ttk.Frame(input_frame)
        files_buttons_frame.grid(row=0, column=1, sticky=(tk.N))
        
        ttk.Button(files_buttons_frame, text="添加文件", 
                  command=self.select_process_files).grid(row=0, column=0, pady=(0, 5))
        ttk.Button(files_buttons_frame, text="移除选中", 
                  command=self.remove_process_file).grid(row=1, column=0, pady=(0, 5))
        ttk.Button(files_buttons_frame, text="清空列表", 
                  command=self.clear_process_files).grid(row=2, column=0)
        
        # 第二行：处理设置（左右分栏）
        # 左侧：数据整合设置
        integrate_frame = ttk.LabelFrame(process_frame, text="2. 数据整合设置", padding="10")
        integrate_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5), pady=(0, 15))
        integrate_frame.columnconfigure(0, weight=1)
        
        # 表头检测选项
        ttk.Label(integrate_frame, text="表头检测策略：", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.header_detection_var = tk.StringVar(value="auto")
        ttk.Radiobutton(integrate_frame, text="自动检测（连续有效行）", 
                       variable=self.header_detection_var, value="auto").grid(row=1, column=0, sticky=tk.W, pady=(0, 2))
        ttk.Radiobutton(integrate_frame, text="首非空行检测", 
                       variable=self.header_detection_var, value="first_non_empty").grid(row=2, column=0, sticky=tk.W, pady=(0, 2))
        
        # 手动指定行号
        manual_row_frame = ttk.Frame(integrate_frame)
        manual_row_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(5, 10))
        manual_row_frame.columnconfigure(1, weight=1)
        
        ttk.Radiobutton(manual_row_frame, text="手动指定起始行：", 
                       variable=self.header_detection_var, value="manual").grid(row=0, column=0, sticky=tk.W)
        
        self.start_row_var = tk.StringVar(value="1")
        start_row_entry = ttk.Entry(manual_row_frame, textvariable=self.start_row_var, width=8)
        start_row_entry.grid(row=0, column=1, sticky=tk.W, padx=(5, 0))
        
        # Sheet处理选项
        ttk.Label(integrate_frame, text="Sheet处理：", font=("Arial", 9, "bold")).grid(row=4, column=0, sticky=tk.W, pady=(10, 5))
        
        self.include_hidden_var = tk.BooleanVar()
        ttk.Checkbutton(integrate_frame, text="包括隐藏Sheet", 
                       variable=self.include_hidden_var).grid(row=5, column=0, sticky=tk.W, pady=(0, 2))
        
        self.merge_sheets_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(integrate_frame, text="合并所有Sheet到单一文件", 
                       variable=self.merge_sheets_var).grid(row=6, column=0, sticky=tk.W, pady=(0, 2))
        
        # 右侧：输出设置和高级选项
        output_frame = ttk.LabelFrame(process_frame, text="3. 输出设置", padding="10")
        output_frame.grid(row=2, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0), pady=(0, 15))
        output_frame.columnconfigure(0, weight=1)
        
        # 输出格式
        ttk.Label(output_frame, text="输出格式：", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.output_format_var = tk.StringVar(value="xlsx")
        format_frame = ttk.Frame(output_frame)
        format_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Radiobutton(format_frame, text="Excel (.xlsx)", 
                       variable=self.output_format_var, value="xlsx").grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(format_frame, text="CSV (.csv)", 
                       variable=self.output_format_var, value="csv").grid(row=0, column=1, sticky=tk.W, padx=(20, 0))
        
        # 高级选项
        ttk.Label(output_frame, text="高级选项：", font=("Arial", 9, "bold")).grid(row=2, column=0, sticky=tk.W, pady=(10, 5))
        
        self.remove_duplicates_var = tk.BooleanVar()
        ttk.Checkbutton(output_frame, text="移除重复行", 
                       variable=self.remove_duplicates_var).grid(row=3, column=0, sticky=tk.W, pady=(0, 2))
        
        self.clean_data_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(output_frame, text="数据清洗（去除空值和异常值）", 
                       variable=self.clean_data_var).grid(row=4, column=0, sticky=tk.W, pady=(0, 2))
        
        self.optimize_memory_var = tk.BooleanVar()
        ttk.Checkbutton(output_frame, text="内存优化（适用于大文件）", 
                       variable=self.optimize_memory_var).grid(row=5, column=0, sticky=tk.W, pady=(0, 10))
        
        # 输出目录选择
        ttk.Label(output_frame, text="输出目录：", font=("Arial", 9, "bold")).grid(row=6, column=0, sticky=tk.W, pady=(0, 5))
        
        output_dir_frame = ttk.Frame(output_frame)
        output_dir_frame.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        output_dir_frame.columnconfigure(0, weight=1)
        
        self.process_output_var = tk.StringVar()
        ttk.Entry(output_dir_frame, textvariable=self.process_output_var, 
                 state="readonly", font=("Arial", 9)).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(output_dir_frame, text="选择", 
                  command=self.select_process_output).grid(row=0, column=1)
        
        # 第三行：执行处理
        action_frame = ttk.Frame(process_frame)
        action_frame.grid(row=3, column=0, columnspan=2, pady=(0, 15))
        
        ttk.Label(action_frame, text="4. 执行处理:", 
                 font=("Arial", 11, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        
        self.process_button = ttk.Button(action_frame, text="开始处理", 
                                        command=self.start_excel_processing, 
                                        style="Accent.TButton")
        self.process_button.grid(row=0, column=1, padx=(0, 10))
        
        ttk.Button(action_frame, text="重置", 
                  command=self.reset_process_settings).grid(row=0, column=2)
        
        # 预览功能按钮（预留）
        ttk.Button(action_frame, text="预览数据", 
                  command=self.preview_data, state="disabled").grid(row=0, column=3, padx=(10, 0))
        
        # 第四行：处理日志
        ttk.Label(process_frame, text="5. 处理日志:", 
                 font=("Arial", 11, "bold")).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        log_frame = ttk.Frame(process_frame)
        log_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.process_log_text = scrolledtext.ScrolledText(log_frame, height=6, font=("Consolas", 9))
        self.process_log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 状态栏
        self.process_status_var = tk.StringVar(value="就绪")
        process_status_bar = ttk.Label(process_frame, textvariable=self.process_status_var, 
                                     relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 9))
        process_status_bar.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 配置行权重
        process_frame.rowconfigure(0, weight=0)  # 功能说明
        process_frame.rowconfigure(1, weight=1)  # 文件输入
        process_frame.rowconfigure(2, weight=2)  # 处理设置
        process_frame.rowconfigure(3, weight=0)  # 执行按钮
        process_frame.rowconfigure(4, weight=0)  # 日志标题
        process_frame.rowconfigure(5, weight=1)  # 日志区域
        process_frame.rowconfigure(6, weight=0)  # 状态栏
        
        # 初始化变量
        self.process_files = []  # 存储选择的文件路径列表
        
        # 初始化日志
        self.process_log("Excel表格处理工具已启动")
        self.process_log("支持智能表头检测、多Sheet处理、数据清洗和格式转换")
    
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
        dialog_frame.rowconfigure(2, weight=1)
        
        # API配置状态显示
        api_status_frame = ttk.Frame(dialog_frame)
        api_status_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        api_status_frame.columnconfigure(1, weight=1)
        
        ttk.Label(api_status_frame, text="API状态:", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky=tk.W)
        self.api_status_label = ttk.Label(api_status_frame, text="", font=("Arial", 9), 
                                        foreground="red")
        self.api_status_label.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        
        # 对话区域
        ttk.Label(dialog_frame, text="快速对话：", font=("Arial", 9)).grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        
        # 对话显示区域
        dialog_display_frame = ttk.Frame(dialog_frame)
        dialog_display_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        dialog_display_frame.columnconfigure(0, weight=1)
        dialog_display_frame.rowconfigure(0, weight=1)
        
        self.dialog_text = scrolledtext.ScrolledText(dialog_display_frame, height=3, font=("Arial", 9), wrap=tk.WORD)
        self.dialog_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 输入和发送
        input_frame = ttk.Frame(dialog_frame)
        input_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(0, weight=1)
        
        self.dialog_input_var = tk.StringVar()
        self.dialog_input = ttk.Entry(input_frame, textvariable=self.dialog_input_var, font=("Arial", 9))
        self.dialog_input.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        self.dialog_input.bind('<Return>', self.send_dialog_message)
        
        ttk.Button(input_frame, text="发送", command=self.send_dialog_message).grid(row=0, column=1)
        
        # 高级对话按钮
        ttk.Button(dialog_frame, text="打开高级对话窗口", 
                  command=self.open_advanced_chat, 
                  style="Accent.TButton").grid(row=5, column=0, pady=(10, 0))
        
        # 输出目录
        ttk.Label(dialog_frame, text="输出目录：", font=("Arial", 9)).grid(row=6, column=0, sticky=tk.W, pady=(10, 5))
        output_controls_frame = ttk.Frame(dialog_frame)
        output_controls_frame.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
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
        self.analysis_log("支持多文件Excel数据脱敏、DeepSeek对话分析和综合报告生成")
        
        # 初始化对话
        self.dialog_text.insert(tk.END, "DeepSeek助手: 您好！我是您的数据分析助手。请先上传Excel文件，然后我们可以开始分析对话。\n\n")
        self.dialog_text.config(state=tk.DISABLED)
    
    def log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_status(self, message, progress=None):
        """更新状态栏和进度条"""
        self.progress_label.config(text=message)
        if progress is not None:
            self.progress_var.set(progress)
        self.root.update_idletasks()
    
    def update_progress(self, percentage, message=""):
        """更新进度条百分比"""
        self.progress_var.set(percentage)
        if message:
            self.progress_label.config(text=f"{message} ({percentage:.1f}%)")
        else:
            self.progress_label.config(text=f"{percentage:.1f}%")
        self.root.update_idletasks()
    
    def get_sheet_names(self, file_path):
        """获取Excel文件的工作表名称列表"""
        try:
            if file_path.endswith('.csv'):
                return ['Sheet1']  # CSV文件只有一个sheet
            
            import openpyxl
            wb = openpyxl.load_workbook(file_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            return sheet_names
        except Exception as e:
            self.log(f"获取工作表名称失败: {str(e)}")
            return []
    
    def on_main_sheet_selected(self, event=None):
        """主表格工作表选择事件"""
        selected_sheet = self.main_sheet_var.get()
        if selected_sheet and hasattr(self.engine, 'main_file_path'):
            # 重新加载主表格数据
            success, message = self.engine.load_main_file(self.engine.main_file_path, selected_sheet)
            if success:
                self.key_column_combo.delete(0, tk.END)
                self.key_column_combo['values'] = self.engine.main_columns
                if self.engine.main_columns:
                    self.key_column_combo.set(self.engine.main_columns[0])
                    self.engine.set_key_column(0)
                self.key_column_info.set(f"主表格共有 {len(self.engine.main_columns)} 列")
                self.log(f"✓ 主表格工作表切换至: {selected_sheet}")
            else:
                self.log(f"✗ 工作表切换失败: {message}")
    
    def on_compare_sheet_selected(self, event=None):
        """对比文件工作表选择事件"""
        selected_sheet = self.compare_sheet_var.get()
        if selected_sheet:
            self.log(f"✓ 对比文件默认工作表设置为: {selected_sheet}")
            # 这里可以添加更多逻辑来处理对比文件的工作表选择
    
    def select_main_file(self):
        """选择主对比表格文件"""
        file_path = filedialog.askopenfilename(
            title="选择主对比表格",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.main_file_var.set(file_path)
            
            # 获取工作表名称列表
            sheet_names = self.get_sheet_names(file_path)
            if sheet_names:
                self.main_sheet_names = sheet_names
                self.main_sheet_combo['values'] = sheet_names
                self.main_sheet_var.set(sheet_names[0])  # 默认选择第一个工作表
                
                # 保存文件路径到引擎
                self.engine.main_file_path = file_path
                
                # 加载主表格
                success, message = self.engine.load_main_table(file_path)
                if success:
                    self.log(f"✓ {message}")
                    
                    # 更新关键列选择
                    self.key_column_combo['values'] = self.engine.main_columns
                    if self.engine.main_columns:
                        self.key_column_combo.set(self.engine.main_columns[0])
                        # 不再自动设置关键列，让用户手动选择
                    
                    # 清空之前的关键列选择
                    self.key_columns.clear()
                    for item in self.key_columns_tree.get_children():
                        self.key_columns_tree.delete(item)
                    
                    self.update_key_column_info()
                    
                    # 更新比对列选择状态
                    self._update_column_selection_status()
            self.log("✓ 已清空所有关键列")
    
    def update_key_column_info(self):
        """更新关键列信息显示"""
        count = len(self.key_columns)
        if count == 0:
            self.key_column_info.set("未选择关键列，请至少选择一个关键列")
        elif count == 1:
            self.key_column_info.set(f"单列匹配: {self.key_columns[0]}")
        else:
            self.key_column_info.set(f"多列联合匹配: {count}个关键列 ({', '.join(self.key_columns[:2])}{'...' if count > 2 else ''})")
            
        # 更新引擎的关键列设置
        self.engine.key_columns = self.key_columns.copy()
    
    def add_key_column(self):
        """添加关键列"""
        selected_column = self.key_column_var.get()
        if selected_column and selected_column in self.engine.main_columns:
            if selected_column not in self.key_columns:
                self.key_columns.append(selected_column)
                # 更新TreeView显示
                self.key_columns_tree.insert('', 'end', values=(selected_column,))
                self.update_key_column_info()
                self.log(f"✓ 添加关键列: {selected_column}")
            else:
                messagebox.showwarning("警告", f"列 '{selected_column}' 已在关键列列表中")
        else:
            messagebox.showwarning("警告", "请先选择有效的列")
    
    def remove_key_column(self):
        """移除选中的关键列"""
        selected_item = self.key_columns_tree.selection()
        if selected_item:
            # 获取选中项的值
            item = self.key_columns_tree.item(selected_item[0])
            column_name = item['values'][0]
            
            # 从列表中移除
            if column_name in self.key_columns:
                self.key_columns.remove(column_name)
                
            # 从TreeView中移除
            self.key_columns_tree.delete(selected_item[0])
            self.update_key_column_info()
            self.log(f"✓ 移除关键列: {column_name}")
        else:
            messagebox.showwarning("警告", "请先选择要移除的关键列")
    
    def clear_key_columns(self):
        """清空所有关键列"""
        if self.key_columns:
            if show_confirm_dialog("确认", "确定要清空所有关键列吗？"):
                self.key_columns.clear()
                # 清空TreeView
                for item in self.key_columns_tree.get_children():
                    self.key_columns_tree.delete(item)
                
                self.update_key_column_info()
                self.log("✓ 已清空所有关键列")

    def _update_column_selection_status(self):
        """更新比对列选择状态信息"""
        if not self.engine.main_table is None and self.compare_files:
            # 获取共同列 - 传递选定的工作表
            selected_sheet = self.compare_sheet_var.get() if hasattr(self, 'compare_sheet_var') else None
            common_columns = self.engine.get_common_columns_from_files(self.compare_files, selected_sheet)
            
            if common_columns:
                if self.engine.selected_columns:
                    selected_count = len(self.engine.selected_columns)
                    total_count = len(common_columns)
                    self.exclude_info.set(f"已选择 {selected_count}/{total_count} 列参与对比")
                else:
                    self.exclude_info.set(f"发现 {len(common_columns)} 个共同列，请选择参与对比的列")
            else:
                self.exclude_info.set("未发现共同列，请检查文件格式")
        else:
            self.exclude_info.set("当前无选中列 - 请先选择主表格和待对比文件")
    
    def on_key_column_selected(self, event=None):
        """关键列选择事件（保持与旧版本兼容）"""
        pass  # 现在使用多关键列功能，这个方法保留以防兼容性问题
    
    def add_compare_files(self):
        """添加待对比文件"""
        file_paths = filedialog.askopenfilenames(
            title="选择待对比文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        
        if file_paths:
            for file_path in file_paths:
                if file_path not in self.compare_files:
                    self.compare_files.append(file_path)
                    # 添加到树形视图
                    filename = os.path.basename(file_path)
                    self.files_tree.insert('', 'end', values=(filename,))
                    
                    # 获取第一个文件的工作表名称用于默认设置
                    if len(self.compare_files) == 1:
                        sheet_names = self.get_sheet_names(file_path)
                        if sheet_names:
                            self.compare_sheet_names = sheet_names
                            self.compare_sheet_combo['values'] = sheet_names
                            self.compare_sheet_var.set(sheet_names[0])
            
            self.log(f"✓ 已添加 {len(file_paths)} 个待对比文件")
            
            # 更新比对列选择状态
            self._update_column_selection_status()
    
    def remove_selected_file(self):
        """移除选中的待对比文件"""
        selected_items = self.files_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要移除的文件")
            return
        
        for item in selected_items:
            # 获取文件名
            filename = self.files_tree.item(item, 'values')[0]
            
            # 从列表中移除对应的文件路径
            for file_path in self.compare_files[:]:
                if os.path.basename(file_path) == filename:
                    self.compare_files.remove(file_path)
                    break
            
            # 从树形视图中移除
            self.files_tree.delete(item)
        
        self.log(f"✓ 已移除 {len(selected_items)} 个文件")
        
        # 更新比对列选择状态
        self._update_column_selection_status()
    
    def clear_files(self):
        self.compare_files.clear()
        # 更新Treeview
        for i in self.files_tree.get_children():
            self.files_tree.delete(i)
        self.log("已清空待对比文件列表")
        self._update_column_selection_status() # 更新列选择按钮状态
    
    def on_smart_match_changed(self):
        """智能匹配选项变化事件"""
        enabled = self.smart_match_var.get()
        success, message = self.engine.set_smart_match(enabled)
        if success:
            self.log(f"✓ {message}")
            
            # 更新比对列选择状态
            self._update_column_selection_status()
        else:
            self.log(f"✗ {message}")
    
    def open_column_selection_window(self):
        """打开列选择窗口"""
        self.log("🔍 尝试打开列选择窗口...")
        
        if self.engine.main_table is None:
            self.log("✗ 未选择主表格")
            messagebox.showwarning("警告", "请先选择主表格")
            return
        
        if not self.compare_files:
            self.log("✗ 未添加待对比文件")
            messagebox.showwarning("警告", "请先添加待对比文件")
            return
        
        self.log(f"📋 主表格已加载: {len(self.engine.main_columns)} 列")
        self.log(f"📋 待对比文件数量: {len(self.compare_files)}")
        
        # 获取共同列 - 传递选定的工作表参数
        selected_sheet = self.compare_sheet_var.get() if hasattr(self, 'compare_sheet_var') else None
        self.log(f"📋 使用工作表: {selected_sheet if selected_sheet else '默认工作表'}")
        
        common_columns = self.engine.get_common_columns_from_files(self.compare_files, selected_sheet)
        self.log(f"📋 发现共同列数量: {len(common_columns)}")
        
        if not common_columns:
            self.log("✗ 未发现共同列")
            messagebox.showwarning("警告", "未发现共同列，请检查文件格式")
            return
        
        self.log("✓ 正在创建列选择窗口...")
        # 创建列选择窗口
        self._show_column_selection_dialog(common_columns)

    def _show_column_selection_dialog(self, common_columns):
        """显示列选择对话框"""
        self.log(f"🎯 创建列选择对话框，共同列: {common_columns[:3]}{'...' if len(common_columns) > 3 else ''}")
        
        try:
            dialog = tk.Toplevel(self.root)
            dialog.title("选择参与对比的列")
            dialog.geometry("600x500")
            dialog.resizable(True, True)
            dialog.transient(self.root)
            dialog.grab_set()
            
            # 居中显示
            dialog.update_idletasks()
            screen_width = dialog.winfo_screenwidth()
            screen_height = dialog.winfo_screenheight()
            window_width = dialog.winfo_width()
            window_height = dialog.winfo_height()
            x = (screen_width // 2) - (window_width // 2)
            y = (screen_height // 2) - (window_height // 2)
            
            # 确保窗口在屏幕范围内
            x = max(0, min(x, screen_width - 600))
            y = max(0, min(y, screen_height - 500))
            
            dialog.geometry(f"600x500+{x}+{y}")
            
            self.log(f"✓ 对话框创建成功 - 屏幕: {screen_width}x{screen_height}, 窗口位置: {x}+{y}")
        except Exception as e:
            self.log(f"✗ 创建对话框失败: {str(e)}")
            messagebox.showerror("错误", f"创建列选择窗口失败: {str(e)}")
            return
        
        try:
            # 配置对话框
            dialog.columnconfigure(0, weight=1)
            dialog.rowconfigure(0, weight=1)
            
            # 主框架
            main_frame = ttk.Frame(dialog, padding="20")
            main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            main_frame.columnconfigure(0, weight=1)
            main_frame.rowconfigure(2, weight=1)
            
            # 标题和说明
            title_label = ttk.Label(main_frame, text="选择参与对比的列", font=("Arial", 14, "bold"))
            title_label.grid(row=0, column=0, pady=(0, 10))
            
            # 获取所有列信息
            main_columns = set(self.engine.main_columns)
            all_compare_columns = set()
            
            # 收集所有对比文件的列
            for file_path in self.compare_files:
                try:
                    excel_file = pd.ExcelFile(file_path)
                    sheet_names = excel_file.sheet_names
                    
                    # 使用选定的工作表，如果指定且存在的话
                    if self.compare_sheet_var.get() and self.compare_sheet_var.get() in sheet_names:
                        first_sheet = self.compare_sheet_var.get()
                    else:
                        first_sheet = sheet_names[0]
                    
                    # 使用智能表头检测
                    header_row, columns = self.engine._find_header_row(file_path, first_sheet)
                    
                    if header_row is not None:
                        # 智能数据类型处理
                        dtype_dict = {}
                        temp_df = pd.read_excel(file_path, sheet_name=first_sheet, header=header_row, nrows=0)
                        for col in temp_df.columns:
                            if self.engine._is_identifier_column(col):
                                dtype_dict[col] = str
                        
                        # 使用检测到的表头行读取数据
                        if dtype_dict:
                            compare_df = pd.read_excel(file_path, sheet_name=first_sheet, header=header_row, dtype=dtype_dict)
                        else:
                            compare_df = pd.read_excel(file_path, sheet_name=first_sheet, header=header_row)
                        all_compare_columns.update(compare_df.columns)
                    else:
                        # 如果智能检测失败，使用默认方式读取
                        dtype_dict = {}
                        temp_df = pd.read_excel(file_path, sheet_name=first_sheet, nrows=0)
                        for col in temp_df.columns:
                            if self.engine._is_identifier_column(col):
                                dtype_dict[col] = str
                        
                        if dtype_dict:
                            compare_df = pd.read_excel(file_path, sheet_name=first_sheet, dtype=dtype_dict)
                        else:
                            compare_df = pd.read_excel(file_path, sheet_name=first_sheet)
                        all_compare_columns.update(compare_df.columns)
                except:
                    continue
            
            # 分类列
            common_cols = main_columns & all_compare_columns
            main_only_cols = main_columns - all_compare_columns
            compare_only_cols = all_compare_columns - main_columns
            
            # 创建所有列的列表（按类别排序）
            all_columns = []
            column_types = {}  # 记录每列的类型
            
            # 添加共同列（蓝色）
            for col in sorted(common_cols):
                all_columns.append(col)
                column_types[col] = 'common'
            
            # 添加主表格独有列（黑色）
            for col in sorted(main_only_cols):
                all_columns.append(col)
                column_types[col] = 'main_only'
            
            # 添加对比文件独有列（灰色）
            for col in sorted(compare_only_cols):
                all_columns.append(col)
                column_types[col] = 'compare_only'
            
            # 更新说明文本
            key_columns_text = ""
            if self.key_columns:
                key_columns_text = f"\n 关键列 {self.key_columns} 为必选项"
            else:
                key_columns_text = "\n 请先选择关键列"
            
            info_text = f"""共发现 {len(all_columns)} 列：
• 蓝色：共同列 ({len(common_cols)} 个) - 可参与对比
• 黑色：主表格独有列 ({len(main_only_cols)} 个) - 可参与对比
• 灰色：对比文件独有列 ({len(compare_only_cols)} 个) - 仅显示，不可选择{key_columns_text}"""
            info_label = ttk.Label(main_frame, text=info_text, font=("Arial", 10), foreground="darkblue")
            info_label.grid(row=1, column=0, pady=(0, 15))
            
            # 列选择区域
            list_frame = ttk.Frame(main_frame)
            list_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
            list_frame.columnconfigure(0, weight=1)
            list_frame.rowconfigure(0, weight=1)
            
            # 创建列表框
            columns_listbox = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, font=("Arial", 10))
            columns_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            
            # 滚动条
            scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=columns_listbox.yview)
            scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
            columns_listbox.configure(yscrollcommand=scrollbar.set)
            
            # 填充列表
            selectable_indices = []  # 记录可选择的索引
            
            for i, col in enumerate(all_columns):
                col_type = column_types[col]
                display_text = col
                
                # 检查是否是关键列
                if col in self.key_columns:
                    display_text += " (关键列 - 必选)"
                
                # 添加类型标识
                if col_type == 'main_only':
                    display_text += " [仅主表格]"
                elif col_type == 'compare_only':
                    display_text += " [仅对比文件]"
                
                columns_listbox.insert(tk.END, display_text)
                
                # 设置颜色
                if col_type == 'common':
                    columns_listbox.itemconfig(i, {'fg': 'blue'})
                    selectable_indices.append(i)
                    # 预选已选择的列
                    if col in self.engine.selected_columns or col in self.key_columns:
                        columns_listbox.selection_set(i)
                elif col_type == 'main_only':
                    columns_listbox.itemconfig(i, {'fg': 'black'})
                    selectable_indices.append(i)  # 主表格独有列也可以选择
                    # 预选已选择的列
                    if col in self.engine.selected_columns or col in self.key_columns:
                        columns_listbox.selection_set(i)
                else:  # compare_only
                    columns_listbox.itemconfig(i, {'fg': 'gray'})
            
            # 操作按钮
            button_frame = ttk.Frame(main_frame)
            button_frame.grid(row=3, column=0, pady=(10, 0))
            
            def select_all():
                # 选择可选择的列（共同列和主表格独有列）
                for idx in selectable_indices:
                    columns_listbox.selection_set(idx)
            
            def clear_selection():
                columns_listbox.selection_clear(0, tk.END)
                # 重新选择关键列
                for i, col in enumerate(all_columns):
                    if col in self.key_columns and i in selectable_indices:
                        columns_listbox.selection_set(i)
            
            def confirm_selection():
                selected_indices = columns_listbox.curselection()
                selected_columns = []
                
                for idx in selected_indices:
                    col = all_columns[idx]
                    # 允许选择共同列和主表格独有列
                    if column_types[col] in ['common', 'main_only']:
                        selected_columns.append(col)
                
                # 确保关键列被选中
                for key_col in self.key_columns:
                    if key_col not in selected_columns:
                        messagebox.showwarning("警告", f"关键列 '{key_col}' 必须被选中")
                        return
                
                # 更新引擎中的选中列
                success, message = self.engine.set_selected_columns(selected_columns)
                if success:
                    self.log(f"✓ {message}")
                    self._update_column_selection_status()
                    dialog.destroy()
                else:
                    self.log(f"✗ {message}")
                    messagebox.showerror("错误", message)
            
            ttk.Button(button_frame, text="全选可用列", command=select_all).grid(row=0, column=0, padx=(0, 10))
            ttk.Button(button_frame, text="清空", command=clear_selection).grid(row=0, column=1, padx=(0, 10))
            ttk.Button(button_frame, text="确认", command=confirm_selection).grid(row=0, column=2, padx=(0, 10))
            ttk.Button(button_frame, text="取消", command=dialog.destroy).grid(row=0, column=3)
            
            self.log("✓ 列选择窗口构建完成")
            
            # 确保窗口正确显示
            dialog.update_idletasks()  # 强制更新窗口布局
            dialog.deiconify()  # 确保窗口不是最小化状态
            dialog.lift()  # 将窗口置于前台
            dialog.focus_force()  # 强制获取焦点
            dialog.attributes('-topmost', True)  # 临时置顶
            dialog.after(100, lambda: dialog.attributes('-topmost', False))  # 100ms后取消置顶
            
            self.log("✓ 列选择窗口已显示并获取焦点")
            
            # 备用显示方法 - 确保窗口可见
            try:
                dialog.wm_state('normal')  # 确保窗口状态正常
                dialog.tkraise()  # 另一种置前方法
                self.root.update()  # 更新主窗口
                self.log("✓ 已执行备用显示方法")
            except Exception as backup_error:
                self.log(f"⚠ 备用显示方法失败: {str(backup_error)}")
            
            # macOS 特殊处理
            import platform
            if platform.system() == 'Darwin':  # macOS
                try:
                    # macOS 特有的窗口显示方法
                    dialog.call('wm', 'attributes', '.', '-topmost', '1')
                    dialog.after(200, lambda: dialog.call('wm', 'attributes', '.', '-topmost', '0'))
                    # 强制刷新
                    dialog.focus_set()
                    self.log("✓ 已应用 macOS 特殊显示方法")
                except Exception as mac_error:
                    self.log(f"⚠ macOS 特殊显示方法失败: {str(mac_error)}")
            
        except Exception as e:
            self.log(f"✗ 构建对话框内容失败: {str(e)}")
            messagebox.showerror("错误", f"构建列选择窗口内容失败: {str(e)}")
            dialog.destroy()

    def select_output_dir(self):
        """选择输出目录"""
        directory = filedialog.askdirectory(title="选择差异报告保存目录")
        if directory:
            self.output_dir_var.set(directory)
            self.log(f"✓ 输出目录设置为: {directory}")
    
    def start_compare(self):
        """开始对比"""
        # 验证输入
        if self.engine.main_table is None:
            messagebox.showwarning("警告", "请先选择主表格")
            return
        
        if not self.compare_files:
            messagebox.showwarning("警告", "请先添加待对比文件")
            return
        
        if not self.key_columns:
            messagebox.showwarning("警告", "请先选择关键列")
            return
        
        if not self.engine.selected_columns:
            messagebox.showwarning("警告", "请先选择参与对比的列")
            return
        
        output_dir = self.output_dir_var.get()
        if not output_dir:
            messagebox.showwarning("警告", "请先选择输出目录")
            return
        
        # 禁用按钮
        self.compare_button.config(state="disabled")
        
        # 重置进度条
        self.update_progress(0, "准备开始对比...")
        
        def progress_callback(message):
            self.log(message)
            # 简单的进度估算
            if "处理文件" in message:
                # 根据文件处理进度更新进度条
                if "[" in message and "/" in message:
                    try:
                        progress_part = message.split("[")[1].split("]")[0]
                        current, total = map(int, progress_part.split("/"))
                        progress = (current / total) * 80  # 文件处理占80%
                        self.update_progress(progress, f"处理文件 {current}/{total}")
                    except:
                        pass
            elif "生成差异报告" in message:
                self.update_progress(90, "生成差异报告...")
            elif "重新排列差异记录" in message:
                self.update_progress(95, "整理报告格式...")
        
        def compare_thread():
            try:
                # 获取用户选择的对比文件工作表
                selected_sheet = self.compare_sheet_var.get() if hasattr(self, 'compare_sheet_var') else None
                
                success, message, output_path = self.engine.compare_tables(
                    self.compare_files, output_dir, progress_callback, selected_sheet
                )
                
                # 在主线程中更新UI
                self.root.after(0, lambda: self._handle_compare_result(success, message, output_path))
                
            except Exception as e:
                error_msg = f"对比过程中发生错误: {str(e)}"
                self.root.after(0, lambda: self._handle_compare_error(error_msg))
        
        # 在后台线程中执行对比
        threading.Thread(target=compare_thread, daemon=True).start()
    
    def _handle_compare_result(self, success, message, output_path):
        """处理对比结果"""
        if success:
            self.update_progress(100, "对比完成")
            self.log(f"✓ {message}")
            
            # 询问是否打开结果文件
            if output_path:
                self.log(f"✓ 结果文件路径: {output_path}")
                try:
                    # 确保对话框在主线程中显示
                    dialog_message = f"对比完成！\n\n{message}\n\n是否打开结果文件所在目录？"
                    if show_confirm_dialog("对比完成", dialog_message):
                        if open_and_highlight_file(output_path):
                            self.log("✓ 已打开结果文件所在目录")
                        else:
                            self.log("⚠ 打开目录失败，请手动打开")
                            messagebox.showinfo("提示", f"对比完成！结果文件保存在:\n{output_path}")
                    else:
                        self.log("用户选择不打开目录")
                except Exception as e:
                    self.log(f"✗ 显示结果对话框时出错: {str(e)}")
                    messagebox.showinfo("对比完成", f"对比完成！\n结果文件保存在:\n{output_path}")
            else:
                self.log("✓ 对比完成，但未生成结果文件")
                messagebox.showinfo("对比完成", "对比完成！未发现任何差异。")
        else:
            self.update_progress(0, "对比失败")
            self.log(f"✗ {message}")
            messagebox.showerror("错误", message)
        
        # 重新启用按钮
        self.compare_button.config(state="normal")
    
    def _handle_compare_error(self, error_msg):
        """处理对比错误"""
        self.update_progress(0, "对比失败")
        self.log(f"✗ {error_msg}")
        messagebox.showerror("错误", error_msg)
        
        # 重新启用按钮
        self.compare_button.config(state="normal")
    
    def reset(self):
        """重置所有设置"""
        self.engine.reset()
        self.log("✓ 已重置所有设置")

    # 表格处理标签页的方法
    def process_log(self, message):
        """添加处理日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.process_log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.process_log_text.see(tk.END)
        self.root.update_idletasks()
    
    def select_process_files(self):
        """选择处理文件"""
        file_paths = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if file_paths:
            for file_path in file_paths:
                if file_path not in self.process_files:
                    self.process_files.append(file_path)
                    filename = os.path.basename(file_path)
                    # 获取工作表数量
                    try:
                        excel_file = pd.ExcelFile(file_path)
                        sheet_count = len(excel_file.sheet_names)
                        status = "待处理"
                    except:
                        sheet_count = "未知"
                        status = "读取失败"
                    
                    self.process_files_tree.insert('', 'end', values=(filename, sheet_count, status))
            
            self.process_log(f"✓ 已添加 {len(file_paths)} 个文件")
    
    def remove_process_file(self):
        """移除选中的处理文件"""
        selected_items = self.process_files_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要移除的文件")
            return
        
        for item in selected_items:
            filename = self.process_files_tree.item(item, 'values')[0]
            for file_path in self.process_files[:]:
                if os.path.basename(file_path) == filename:
                    self.process_files.remove(file_path)
                    break
            self.process_files_tree.delete(item)
        
        self.process_log(f"✓ 已移除 {len(selected_items)} 个文件")
    
    def clear_process_files(self):
        """清空处理文件列表"""
        if self.process_files:
            if show_confirm_dialog("确认", "确定要清空所有文件吗？"):
                self.process_files.clear()
                for item in self.process_files_tree.get_children():
                    self.process_files_tree.delete(item)
                self.process_log("✓ 已清空文件列表")
    
    def select_process_output(self):
        """选择处理输出目录"""
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.process_output_var.set(directory)
            self.process_log(f"✓ 输出目录设置为: {directory}")
    
    def start_excel_processing(self):
        """开始Excel处理"""
        messagebox.showinfo("提示", "表格处理功能正在开发中，敬请期待！")
    
    def reset_process_settings(self):
        """重置处理设置"""
        if show_confirm_dialog("确认", "确定要重置所有设置吗？"):
            self.process_files.clear()
            for item in self.process_files_tree.get_children():
                self.process_files_tree.delete(item)
            self.process_output_var.set("")
            self.header_detection_var.set("auto")
            self.start_row_var.set("1")
            self.include_hidden_var.set(False)
            self.merge_sheets_var.set(True)
            self.output_format_var.set("xlsx")
            self.remove_duplicates_var.set(False)
            self.clean_data_var.set(True)
            self.optimize_memory_var.set(False)
            self.process_log("✓ 已重置所有设置")
    
    def preview_data(self):
        """预览数据"""
        messagebox.showinfo("提示", "数据预览功能正在开发中，敬请期待！")
    
    # 数据分析标签页的方法
    def analysis_log(self, message):
        """添加分析日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.analysis_log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.analysis_log_text.see(tk.END)
        self.root.update_idletasks()
    
    def select_analysis_file(self):
        """选择分析文件"""
        file_paths = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        
        if file_paths:
            for file_path in file_paths:
                if file_path not in self.analysis_files:
                    self.analysis_files.append(file_path)
                    filename = os.path.basename(file_path)
                    status = "已添加"
                    self.analysis_files_tree.insert('', 'end', values=(filename, status))
            
            self.analysis_log(f"✓ 已添加 {len(file_paths)} 个分析文件")
    
    def remove_analysis_file(self):
        """移除选中的分析文件"""
        selected_items = self.analysis_files_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要移除的文件")
            return
        
        for item in selected_items:
            filename = self.analysis_files_tree.item(item, 'values')[0]
            for file_path in self.analysis_files[:]:
                if os.path.basename(file_path) == filename:
                    self.analysis_files.remove(file_path)
                    break
            self.analysis_files_tree.delete(item)
        
        self.analysis_log(f"✓ 已移除 {len(selected_items)} 个文件")
    
    def clear_analysis_files(self):
        """清空分析文件列表"""
        if self.analysis_files:
            if show_confirm_dialog("确认", "确定要清空所有文件吗？"):
                self.analysis_files.clear()
                for item in self.analysis_files_tree.get_children():
                    self.analysis_files_tree.delete(item)
                self.analysis_log("✓ 已清空文件列表")
    
    def select_analysis_output(self):
        """选择分析输出目录"""
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.analysis_output_var.set(directory)
            self.analysis_log(f"✓ 输出目录设置为: {directory}")
    
    def send_dialog_message(self, event=None):
        """发送对话消息"""
        messagebox.showinfo("提示", "对话功能正在开发中，敬请期待！")
    
    def open_advanced_chat(self):
        """打开高级对话窗口"""
        messagebox.showinfo("提示", "高级对话功能正在开发中，敬请期待！")
    
    def start_analysis(self):
        """开始分析"""
        messagebox.showinfo("提示", "数据分析功能正在开发中，敬请期待！")
    
    def reset_analysis(self):
        """重置分析设置"""
        if show_confirm_dialog("确认", "确定要重置所有设置吗？"):
            self.analysis_files.clear()
            for item in self.analysis_files_tree.get_children():
                self.analysis_files_tree.delete(item)
            self.analysis_output_var.set("")
            self.sensitive_listbox.selection_clear(0, tk.END)
            self.sensitive_row_listbox.selection_clear(0, tk.END)
            self.row_select_mode.set("none")
            self.dialog_input_var.set("")
            self.analysis_log("✓ 已重置所有设置")

if __name__ == "__main__":
    root = tk.Tk()
    app = ImprovedExcelCompareGUI(root)
    root.mainloop()
