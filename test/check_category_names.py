import pandas as pd

# 读取参考文件
reference_file = '/Users/monokeros/Project/Coding/CopySelector/test/参考文件.xlsx'
df = pd.read_excel(reference_file, sheet_name='Sheet1')

# 提取B类产品列
print("B类产品列名:")
print(df.columns[5:])  # 从第6列开始（0-based）

# 提取第一行的产品名称（B类）
print("\nB类产品名称:")
b_category_names = df.iloc[0, 5:].tolist()
print(b_category_names)

# 读取结果文件的sheet页
result_file = '/Users/monokeros/Project/Coding/CopySelector/test/结果文件.xlsx'
xls = pd.ExcelFile(result_file)
print("\n结果文件sheet页:")
print(xls.sheet_names)
