import os
import json
from excel_processor import ExcelProcessor

class ExcelSelector:
    def __init__(self, config_file='config.json'):
        # 配置文件路径
        self.config_file = config_file
        
        # 加载配置
        if os.path.exists(self.config_file):
            print("Loading configuration from file...")
            self.config = self.load_config()
        else:
            raise FileNotFoundError(f"Configuration file not found: {self.config_file}")
        
        # 初始化Excel处理器
        self.processor = ExcelProcessor()
    
    def load_config(self):
        """加载配置文件"""
        with open(self.config_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def save_config(self):
        """保存配置文件"""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)
    
    def read_source_file(self):
        """读取参考文件，提取产品的勾选状态"""
        return self.processor.read_source_file(
            self.config['source_file'],
            self.config['source_config']['category_row'],
            self.config['source_config']['indicator_start_row']
        )
    
    def read_target_file(self):
        """读取结果文件，获取sheet页信息"""
        return self.processor.read_target_file(self.config['target_file'])
    
    def map_data(self, source_status):
        """将参考文件的勾选状态映射到结果文件"""
        self.processor.map_data(
            source_status,
            self.config['target_file'],
            self.config['category_mapping'],
            self.config['month_config'],
            self.config['check_marks'],
            self.config['target_config']['indicator_start_row']
        )
    
    def run(self):
        """运行整个映射过程"""
        print("Reading source file...")
        source_status = self.read_source_file()
        print(f"Source status keys: {list(source_status.keys())}")
        print(f"Category mapping: {self.config['category_mapping']}")
        
        print("Reading target file...")
        target_sheets = self.read_target_file()
        print(f"Target sheets: {target_sheets}")
        
        print("Mapping data...")
        self.map_data(source_status)
        
        # 保存配置
        self.save_config()
        print("Configuration saved successfully!")
        
        print("Process completed successfully!")

if __name__ == "__main__":
    selector = ExcelSelector()
    selector.run()
