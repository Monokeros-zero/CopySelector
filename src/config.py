#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
配置管理模块
"""

import os
import json
import tempfile

class ConfigManager:
    """配置管理类"""
    
    def __init__(self):
        """初始化配置管理器"""
        self.config_dir = os.path.join(os.path.dirname(__file__), "..", "configs")
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)
        
    def get_config_files(self):
        """获取所有配置文件"""
        config_files = []
        for file in os.listdir(self.config_dir):
            if file.endswith('.json'):
                config_files.append(file)
        return config_files
    
    def get_last_config(self):
        """获取最后修改的配置文件"""
        config_files = self.get_config_files()
        if config_files:
            # 按修改时间排序，选择最新的
            config_files.sort(key=lambda x: os.path.getmtime(os.path.join(self.config_dir, x)), reverse=True)
            return config_files[0]
        return None
    
    def load_config(self, config_name):
        """加载配置文件"""
        config_file = os.path.join(self.config_dir, config_name)
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"加载配置失败: {str(e)}")
        return None
    
    def save_config(self, config_name, config):
        """保存配置文件"""
        config_file = os.path.join(self.config_dir, config_name)
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存配置失败: {str(e)}")
            return False
    
    def create_temp_config(self, config):
        """创建临时配置文件"""
        temp_config = tempfile.NamedTemporaryFile(suffix='.json', delete=False)
        with open(temp_config.name, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return temp_config.name
    
    def delete_temp_config(self, temp_config_path):
        """删除临时配置文件"""
        if os.path.exists(temp_config_path):
            os.unlink(temp_config_path)
