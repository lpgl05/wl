import sys
sys.path.insert(0, r'D:\code\vscode\wl')

from main import extract_sample_and_concentration

if __name__ == "__main__":
    print("\n测试挥发酚文件:")
    file_path = r"D:\code\vscode\wl\data\0953-DX-挥发酚.xlsx"
    result = extract_sample_and_concentration(file_path, skip_empty_rows=True)
