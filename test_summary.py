import sys
import os
import re
sys.path.insert(0, r'D:\code\vscode\wl')

from main import get_excel_files_from_folder, extract_sample_and_concentration
import pandas as pd

if __name__ == "__main__":
    folder_path = r"D:\code\vscode\wl\data"
    excel_files = get_excel_files_from_folder(folder_path)
    
    print("处理Excel文件并汇总:")
    print("="*80)
    
    # 创建一个总的数组来存储所有项目的数据
    all_data = []
    
    for excel_file in excel_files:
        result_array = extract_sample_and_concentration(excel_file, skip_empty_rows=True)
        
        # 获取文件名并提取汉字部分
        file_name = os.path.basename(excel_file)
        chinese_chars = re.findall(r'[\u4e00-\u9fff]+', file_name)
        chinese_name = ''.join(chinese_chars) if chinese_chars else file_name
        
        # 基于result_array,再新建一列（第一列），内容为chinese_name
        for row in result_array:
            row.insert(0, chinese_name)
            all_data.append(row)
        
        print(f"✓ {chinese_name}: 提取了 {len(result_array)} 行")
    
    # 将所有数据写入到一个Excel文件中
    if all_data:
        output_file = os.path.join(folder_path, "汇总_提取结果.xlsx")
        df = pd.DataFrame(all_data, columns=["项目名", "样品编号", "浓度"])
        df.to_excel(output_file, index=False)
        
        print(f"\n{'='*80}")
        print(f"✅ 所有提取结果已汇总保存到: {output_file}")
        print(f"   共提取 {len(all_data)} 行数据，来自 {len(excel_files)} 个文件")
        print(f"{'='*80}")
        
        # 显示汇总结果预览
        print("\n汇总数据预览 (前10行):")
        print(df.head(10).to_string(index=False))
        
        print(f"\n按项目分组统计:")
        print(df.groupby('项目名').size())
    else:
        print("\n⚠ 没有提取到任何数据")
