import pandas as pd
import os

# 读取参考文件
reference_file = '/Users/monokeros/Project/Coding/CopySelector/test/参考文件.xlsx'
result_file = '/Users/monokeros/Project/Coding/CopySelector/test/结果文件.xlsx'

print("=== 参考文件信息 ===")
# 获取参考文件的sheet页
reference_xls = pd.ExcelFile(reference_file)
print(f"参考文件sheet页: {reference_xls.sheet_names}")

# 读取sheet1的内容
sheet1_df = pd.read_excel(reference_file, sheet_name='Sheet1')
print("\nSheet1内容:")
print(sheet1_df)
print(f"\nSheet1形状: {sheet1_df.shape}")

print("\n=== 结果文件信息 ===")
# 获取结果文件的sheet页
result_xls = pd.ExcelFile(result_file)
print(f"结果文件sheet页: {result_xls.sheet_names}")

# 读取第一个sheet的内容
result_sheet1 = pd.read_excel(result_file, sheet_name=result_xls.sheet_names[0])
print(f"\n{result_xls.sheet_names[0]}内容:")
print(result_sheet1)
print(f"\n{result_xls.sheet_names[0]}形状: {result_sheet1.shape}")
