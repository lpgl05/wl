"""
简化测试版本 - 用于诊断打包问题
"""
import sys
print("Step 1: 程序启动...")

try:
    import tkinter
    print("Step 2: tkinter 导入成功")
except Exception as e:
    print(f"Step 2 失败: tkinter 导入错误 - {e}")
    sys.exit(1)

try:
    from tkinter import filedialog
    print("Step 3: filedialog 导入成功")
except Exception as e:
    print(f"Step 3 失败: filedialog 导入错误 - {e}")
    sys.exit(1)

try:
    import pandas as pd
    print("Step 4: pandas 导入成功")
except Exception as e:
    print(f"Step 4 失败: pandas 导入错误 - {e}")
    sys.exit(1)

try:
    from openpyxl import load_workbook
    print("Step 5: openpyxl 导入成功")
except Exception as e:
    print(f"Step 5 失败: openpyxl 导入错误 - {e}")
    sys.exit(1)

print("\n所有模块导入成功!")
print("准备显示文件夹选择对话框...")

try:
    root = tkinter.Tk()
    root.withdraw()
    print("Step 6: Tk 窗口创建成功")
    
    folder_path = filedialog.askdirectory(title="测试 - 选择文件夹")
    print(f"Step 7: 用户选择的文件夹: {folder_path}")
    
    if folder_path:
        print(f"✅ 成功选择文件夹: {folder_path}")
    else:
        print("⚠️ 用户取消了选择")
        
except Exception as e:
    print(f"对话框显示失败: {e}")
    import traceback
    traceback.print_exc()

print("\n按回车键退出...")
input()
