import json
import os
import sys

# 添加src目录到路径
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

from excel_selector import ExcelSelector

def create_test_config():
    """创建测试配置文件"""
    config = {
        'source_file': os.path.join(os.path.dirname(__file__), '参考文件.xlsx'),
        'target_file': os.path.join(os.path.dirname(__file__), '结果文件.xlsx'),
        'source_config': {
            'indicator_position': 'column',
            'indicator_start_row': 3,  # 从第3行开始（Excel行号）
            'category_row': 1,  # 产品行在第1行（Excel行号）
            'b_category_start_col': 5  # B类产品从第6列开始（0-based）
        },
        'target_config': {
            'indicator_position': 'column',
            'indicator_start_row': 3,  # 从第3行开始（Excel行号）
            'month_row': 2  # 月份行在第2行（Excel行号）
        },
        'check_marks': {
            'source': ['✓'],
            'target': '✓'
        },
        'category_mapping': {
            '产品B1': '产品B1',
            '产品B2': '产品B2',
            '产品B3': '产品B3'
        },
        'month_config': {
            'selected_months': [3, 4, 5, 6, 7, 8, 9, 10],  # 3-10月
            'month_column_mapping': {}
        }
    }
    
    # 保存配置到测试目录
    config_file = os.path.join(os.path.dirname(__file__), 'test_config.json')
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    
    return config_file

def run_test():
    """运行测试"""
    # 创建测试配置
    config_file = create_test_config()
    print(f"Created test configuration: {config_file}")
    
    # 运行ExcelSelector
    try:
        selector = ExcelSelector(config_file=config_file)
        selector.run()
        print("Test completed successfully!")
    except Exception as e:
        print(f"Error during test: {e}")

if __name__ == "__main__":
    run_test()
