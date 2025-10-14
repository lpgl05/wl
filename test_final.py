import sys
sys.path.insert(0, r'D:\code\vscode\wl')

from main import extract_sample_and_concentration

if __name__ == "__main__":
    files = {
        "总硬度": r"D:\code\vscode\wl\data\0953-DX-总硬度.xlsx",
        "挥发酚": r"D:\code\vscode\wl\data\0953-DX-挥发酚.xlsx",
        "氰化物": r"D:\code\vscode\wl\data\0953-DX-氰化物.xlsx"
    }
    
    for name, file_path in files.items():
        print(f"\n{'#'*80}")
        print(f"# 测试{name}文件")
        print(f"{'#'*80}")
        result = extract_sample_and_concentration(file_path, skip_empty_rows=True)
        print(f"\n✅ {name}: 提取成功 {len(result)} 行数据")
