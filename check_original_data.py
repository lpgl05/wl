import pandas as pd

file_path = r"D:\code\vscode\wl\data\0953-DX-挥发酚.xlsx"
df = pd.read_excel(file_path, header=2)

print("原始数据中的样品编号和浓度:")
print("="*80)
for i in range(min(10, len(df))):
    sample = df.iloc[i]['编号']
    conc = df.iloc[i]['浓度mg/L']
    if pd.notna(sample):
        print(f"行 {i}: 编号='{sample}', 浓度={conc}, 有平行={'平行' in str(sample)}")
