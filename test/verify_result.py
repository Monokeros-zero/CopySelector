import pandas as pd

# 读取结果文件
result_file = '/Users/monokeros/Project/Coding/CopySelector/test/结果文件.xlsx'
xls = pd.ExcelFile(result_file)

# 验证每个sheet页
for sheet_name in xls.sheet_names:
    print(f"\n=== 验证 {sheet_name} ===")
    df = pd.read_excel(result_file, sheet_name=sheet_name)
    
    # 提取指标行（从第3行开始，Excel行号）
    indicator_start_idx = 2  # 0-based
    indicators = df.iloc[indicator_start_idx:, 0].tolist()
    
    # 检查3-10月的勾选情况
    month_cols = ['3月预测', '4月预测', '5月预测', '6月预测', '7月预测', '8月预测', '9月预测', '10月预测']
    
    # 统计勾选的指标数量
    checked_count = 0
    for i, indicator in enumerate(indicators):
        # 检查是否在所有月份都被勾选
        all_checked = True
        for month_col in month_cols:
            if month_col in df.columns:
                cell_value = df.iloc[indicator_start_idx + i, df.columns.get_loc(month_col)]
                if cell_value != '✓':
                    all_checked = False
                    break
        if all_checked:
            checked_count += 1
            print(f"指标 {indicator}: 已勾选")
    
    print(f"\n{sheet_name} 中已勾选的指标数量: {checked_count}")

print("\n验证完成！")
