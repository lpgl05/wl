import pandas as pd

# 读取生成的汇总文件
file_path = r"D:\code\vscode\wl\data\汇总_提取结果.xlsx"
df = pd.read_excel(file_path)

print("汇总Excel文件内容:")
print("="*80)
print(f"总行数: {len(df)}")
print(f"列名: {list(df.columns)}")
print("\n完整数据:")
print(df.to_string(index=False))

print("\n"+"="*80)
print("统计信息:")
print("="*80)
print(f"\n各项目数据量:")
print(df.groupby('项目名').size())

print(f"\n各样品编号数据量:")
print(df.groupby('样品编号').size())
