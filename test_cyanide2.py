import sys
sys.path.insert(0, r'D:\code\vscode\wl')

from main import extract_sample_and_concentration

file_path = r"D:\code\vscode\wl\data\0953-DX氰化物.xlsx"
result = extract_sample_and_concentration(file_path, skip_empty_rows=True)

print("\n完整数据:")
for i, row in enumerate(result):
    print(f"  行{i}: {row}")
