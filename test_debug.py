import sys
sys.path.insert(0, r'D:\code\vscode\wl')
import pandas as pd

from main import extract_sample_and_concentration

if __name__ == "__main__":
    file_path = r"D:\code\vscode\wl\data\0953-DX-挥发酚.xlsx"
    result = extract_sample_and_concentration(file_path, skip_empty_rows=True)
    
    print("\n完整result_array:")
    for i, row in enumerate(result):
        if i <= 10:
            print(f"行 {i}: {row}")
