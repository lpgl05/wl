import pandas as pd

file_path = r"D:\code\vscode\wl\data\0953-DX氰化物.xlsx"

print("测试读取氰化物文件...")
try:
    df = pd.read_excel(file_path, header=1)
    print(f"成功读取，列名: {list(df.columns[:10])}")
    print(f"数据行数: {len(df)}")
    
    sample_col_names = ['样品编号', '样品名称', '样品', '编号']
    concentration_col_names = ['浓度mg/L', '浓度mg/l', '浓度']
    
    found_sample = None
    for col in sample_col_names:
        if col in df.columns:
            found_sample = col
            break
    
    found_conc = None
    for col in concentration_col_names:
        if col in df.columns:
            found_conc = col
            break
    
    print(f"找到样品列: {found_sample}")
    print(f"找到浓度列: {found_conc}")
except Exception as e:
    print(f"出错: {e}")
    import traceback
    traceback.print_exc()
