import sys
import os
import re
sys.path.insert(0, r'D:\code\vscode\wl')

from main import get_excel_files_from_folder, extract_sample_and_concentration

if __name__ == "__main__":
    folder_path = r"D:\code\vscode\wl\data"
    excel_files = get_excel_files_from_folder(folder_path)
    
    print("处理Excel文件:")
    print("="*80)
    
    for excel_file in excel_files:
        # 获取文件名并提取汉字部分
        file_name = os.path.basename(excel_file)
        chinese_chars = re.findall(r'[\u4e00-\u9fff]+', file_name)
        chinese_name = ''.join(chinese_chars) if chinese_chars else file_name
        
        print(f"\n文件: {file_name}")
        print(f"提取的汉字: {chinese_name}")
        
        result_array = extract_sample_and_concentration(excel_file, skip_empty_rows=True)
        print(f"✅ 提取了 {len(result_array)} 行符合条件的数据")
        print("-"*80)
