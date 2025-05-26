#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel工具环境配置管理模块
用于管理DeepSeek API Key和其他配置参数
"""

import os
import sys
from pathlib import Path

class Config:
    """配置管理类"""
    
    def __init__(self):
        self.config_file = ".env"
        self.config_data = {}
        self.load_config()
    
    def load_config(self):
        """加载配置"""
        # 默认配置
        self.config_data = {
            'DEEPSEEK_API_KEY': '',
            'DEEPSEEK_BASE_URL': 'https://api.deepseek.com',
            'DEEPSEEK_DEFAULT_MODEL': 'deepseek-chat',
            'DEEPSEEK_TIMEOUT': '30',
            'DEEPSEEK_DEBUG': 'false'
        }
        
        # 尝试从环境变量读取
        for key in self.config_data.keys():
            env_value = os.getenv(key)
            if env_value:
                self.config_data[key] = env_value
        
        # 尝试从.env文件读取
        env_file = Path(self.config_file)
        if env_file.exists():
            try:
                with open(env_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith('#') and '=' in line:
                            key, value = line.split('=', 1)
                            self.config_data[key.strip()] = value.strip()
            except Exception as e:
                print(f"警告: 读取.env文件失败: {e}")
    
    def save_config(self):
        """保存配置到.env文件"""
        try:
            env_content = """# DeepSeek API 配置文件
# 请将此文件添加到 .gitignore 中以确保API Key不会被提交到版本控制系统

# DeepSeek API Key
DEEPSEEK_API_KEY={api_key}

# DeepSeek API 基础URL（可选，默认为官方地址）
DEEPSEEK_BASE_URL={base_url}

# 默认模型设置（可选）
DEEPSEEK_DEFAULT_MODEL={default_model}

# API调用超时设置（秒）
DEEPSEEK_TIMEOUT={timeout}

# 是否启用调试模式
DEEPSEEK_DEBUG={debug}
""".format(
                api_key=self.config_data.get('DEEPSEEK_API_KEY', ''),
                base_url=self.config_data.get('DEEPSEEK_BASE_URL', 'https://api.deepseek.com'),
                default_model=self.config_data.get('DEEPSEEK_DEFAULT_MODEL', 'deepseek-chat'),
                timeout=self.config_data.get('DEEPSEEK_TIMEOUT', '30'),
                debug=self.config_data.get('DEEPSEEK_DEBUG', 'false')
            )
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                f.write(env_content)
            
            print(f"✓ 配置已保存到 {self.config_file}")
            return True
        except Exception as e:
            print(f"✗ 保存配置失败: {e}")
            return False
    
    def get(self, key, default=None):
        """获取配置值"""
        return self.config_data.get(key, default)
    
    def set(self, key, value):
        """设置配置值"""
        self.config_data[key] = str(value)
    
    def get_api_key(self):
        """获取API Key"""
        return self.get('DEEPSEEK_API_KEY', '')
    
    def set_api_key(self, api_key):
        """设置API Key"""
        self.set('DEEPSEEK_API_KEY', api_key)
    
    def get_base_url(self):
        """获取基础URL"""
        return self.get('DEEPSEEK_BASE_URL', 'https://api.deepseek.com')
    
    def get_default_model(self):
        """获取默认模型"""
        return self.get('DEEPSEEK_DEFAULT_MODEL', 'deepseek-chat')
    
    def get_timeout(self):
        """获取超时设置"""
        try:
            return int(self.get('DEEPSEEK_TIMEOUT', '30'))
        except ValueError:
            return 30
    
    def is_debug_enabled(self):
        """是否启用调试模式"""
        return self.get('DEEPSEEK_DEBUG', 'false').lower() in ('true', '1', 'yes', 'on')
    
    def validate_api_key(self):
        """验证API Key格式"""
        api_key = self.get_api_key()
        if not api_key:
            return False, "API Key为空"
        
        if not api_key.startswith('sk-'):
            return False, "API Key格式错误，应该以'sk-'开头"
        
        if len(api_key) < 20:
            return False, "API Key长度太短"
        
        return True, "API Key格式正确"
    
    def create_env_file(self):
        """创建.env文件"""
        return self.save_config()
    
    def update_gitignore(self):
        """更新.gitignore文件以忽略.env"""
        gitignore_path = Path('.gitignore')
        gitignore_content = ""
        
        # 读取现有.gitignore内容
        if gitignore_path.exists():
            try:
                with open(gitignore_path, 'r', encoding='utf-8') as f:
                    gitignore_content = f.read()
            except Exception as e:
                print(f"警告: 读取.gitignore失败: {e}")
        
        # 检查是否已经包含.env
        if '.env' not in gitignore_content:
            # 添加.env到.gitignore
            if gitignore_content and not gitignore_content.endswith('\n'):
                gitignore_content += '\n'
            
            gitignore_content += """
# 环境配置文件（包含敏感信息）
.env
config.local.*
*.key
"""
            
            try:
                with open(gitignore_path, 'w', encoding='utf-8') as f:
                    f.write(gitignore_content)
                print("✓ 已更新.gitignore文件")
                return True
            except Exception as e:
                print(f"✗ 更新.gitignore失败: {e}")
                return False
        else:
            print("✓ .gitignore已包含.env规则")
            return True

# 全局配置实例
config = Config()

def get_api_key():
    """快速获取API Key的便捷函数"""
    return config.get_api_key()

def set_api_key(api_key):
    """快速设置API Key的便捷函数"""
    config.set_api_key(api_key)
    return config.save_config()

def validate_and_setup_config():
    """验证并设置配置"""
    print("=== DeepSeek API 配置验证 ===")
    
    # 验证API Key
    is_valid, message = config.validate_api_key()
    if is_valid:
        print(f"✓ {message}")
        print(f"  API Key: {config.get_api_key()[:10]}...")
    else:
        print(f"✗ {message}")
        print(f"  当前API Key: {config.get_api_key()}")
    
    # 显示其他配置
    print(f"  基础URL: {config.get_base_url()}")
    print(f"  默认模型: {config.get_default_model()}")
    print(f"  超时设置: {config.get_timeout()}秒")
    print(f"  调试模式: {'开启' if config.is_debug_enabled() else '关闭'}")
    
    # 创建.env文件
    if config.create_env_file():
        print("✓ .env文件已创建/更新")
    
    # 更新.gitignore
    config.update_gitignore()
    
    return is_valid

if __name__ == "__main__":
    # 测试配置管理功能
    validate_and_setup_config() 