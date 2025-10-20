import os
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
from datetime import datetime
from tkinter import Tk, messagebox, simpledialog
from tkinter.filedialog import askdirectory, asksaveasfilename

# 创建一个函数，让用户打开指定文件夹，遍历该文件夹中的所有excel文件
# 文件夹由用户指定，弹出一个文件管理器，让用户自己去选择文件夹
def select_folder():
    root = Tk()
    root.withdraw()  # we don't want a full GUI, so keep the root window from appearing
    folder_path = askdirectory(title="Select Folder Containing Excel Files")
    root.destroy()
    return folder_path

# 在遍历文件夹中的excel文件时，应该去掉 类似这种的临时文件 ~$0953-DX-挥发酚.xlsx
def get_excel_files_from_folder(folder_path):
    excel_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xls'):
                if not file.startswith('~$'):
                    excel_files.append(os.path.join(root, file))
    return excel_files

def _normalize_sample_id(value: object) -> str:
    """
    将样品编号做规范化，便于匹配：
    - 转大写
    - 去除空白、全角空白
    - 去除“平行X”字样
    - 去除中文括号及常见标点
    """
    s = str(value) if value is not None else ""
    s = s.upper().strip()
    # 去掉空白
    s = re.sub(r"\s+", "", s)
    # 去掉（平行1/2/...）字样
    s = re.sub(r"平行[0-9]+", "", s)
    # 去掉中英文括号及其中的空内容痕迹
    s = s.replace("（", "(").replace("）", ")").replace("【", "[").replace("】", "]")
    s = s.replace("，", ",").replace("。", ".")
    return s

def extract_sample_and_concentration(file_path, skip_empty_rows=False, targets=None):
    """
    提取Excel文件中的样品编号和浓度列数据
    
    参数:
        file_path: Excel文件的完整路径
        skip_empty_rows: 是否跳过样品编号和浓度都为空的行 (默认False)
        targets: 可选，目标样品编号列表（例如 ["DX2509530201", "DX2509530301"]）。
                 若提供，则只返回匹配这些样品编号的行。
    
    返回:
        二维数组，每行包含[样品编号, 浓度]
    """
    print(f"\n{'='*80}")
    print(f"处理文件: {file_path}")
    print(f"{'='*80}")
    
    # 样品编号和浓度列名的可能变体
    sample_col_names = ['样品编号', '样品名称', '样品', '编号']
    concentration_col_names = ['浓度mg/L', '浓度mg/l', '浓度']
    
    df = None
    header_row = None
    sample_col = None
    concentration_col = None
    
    # 尝试不同的header行 (0-99，即第1-100行) 来找到包含目标列的行
    for header in range(100):
        try:
            temp_df = pd.read_excel(file_path, header=header)
            
            # 查找样品编号列
            temp_sample_col = None
            for col_name in sample_col_names:
                if col_name in temp_df.columns:
                    temp_sample_col = col_name
                    break
            
            # 查找浓度列
            temp_concentration_col = None
            for col_name in concentration_col_names:
                if col_name in temp_df.columns:
                    temp_concentration_col = col_name
                    break
            
            # 如果找到了两个目标列，就使用这个header
            if temp_sample_col and temp_concentration_col:
                df = temp_df
                header_row = header
                sample_col = temp_sample_col
                concentration_col = temp_concentration_col
                print(f"✓ 使用第 {header+1} 行作为列名 (header={header})")
                break
        except Exception as e:
            continue
    
    # 如果没有找到合适的列
    if df is None or sample_col is None or concentration_col is None:
        print(f"\n⚠ 警告: 未找到样品编号或浓度列")
        if df is not None:
            print(f"可用的列名: {list(df.columns)}")
        return []
    
    print(f"✓ 找到样品编号列: '{sample_col}'")
    print(f"✓ 找到浓度列: '{concentration_col}'")
    
    # 提取两列数据
    sample_data = df[sample_col].tolist()
    concentration_data = df[concentration_col].tolist()
    
    # 组合成二维数组
    result_array = []
    for sample, concentration in zip(sample_data, concentration_data):
        # 如果需要跳过空行
        if skip_empty_rows:
            # 检查是否两个值都为空
            if pd.isna(sample) and pd.isna(concentration):
                continue
        result_array.append([sample, concentration])
    
    # 第一步：填充平行样品中浓度为nan的行
    for i, row in enumerate(result_array):
        # 根据样品编号决定是否需要处理
        if '250953' in str(row[0]) and 'KB' not in str(row[0]) and 'PS' not in str(row[0]):
            # 如果样品编号中存在 "平行"， 且对应的浓度列为NaN时， 则用前后行的浓度值进行填充
            if '平行' in str(row[0]) and pd.isna(row[1]):
                # 向前查找有浓度值的平行样品
                for j in range(i-1, -1, -1):
                    if '平行' in str(result_array[j][0]) and not pd.isna(result_array[j][1]):
                        row[1] = result_array[j][1]
                        break
                # 如果向前没找到，向后查找有浓度值的平行样品
                if pd.isna(row[1]):
                    for j in range(i+1, len(result_array)):
                        if '平行' in str(result_array[j][0]) and not pd.isna(result_array[j][1]):
                            row[1] = result_array[j][1]
                            break
    
    # 第二步：去掉所有"平行1"、"平行2"等字样
    #for i, row in enumerate(result_array):
    #    if '250953' in str(row[0]) and 'KB' not in str(row[0]) and 'PS' not in str(row[0]):
    #        if '平行' in str(row[0]):
    #            row[0] = re.sub(r'平行[1-9]+', '', str(row[0]))
    
    # 筛选符合条件的行，形成新数组
    filtered_array = result_array.copy()
    #for i, row in enumerate(result_array):
    #    # 只保留包含250953且不包含KB、PS的行
    #    if '250953' in str(row[0]) and 'KB' not in str(row[0]) and 'PS' not in str(row[0]):
    #        filtered_array.append(row)

    # 如果指定了目标样品编号，则进一步按目标过滤（按规范化后精确匹配）
    if targets:
        normalized_targets = {_normalize_sample_id(t) for t in targets}
        filtered_array = [
            [row[0], row[1]]
            for row in filtered_array
            if _normalize_sample_id(row[0]) in normalized_targets
        ]
    
    # 打印结果
    print(f"\n提取的二维数组 (共 {len(result_array)} 行):")
    if skip_empty_rows:
        print("(已过滤掉空行)")
    print(f"列顺序: [样品编号, 浓度]")
    
    print(f"\n符合条件的数据 (共 {len(filtered_array)} 行):")
    print("(包含250953 且 不包含KB、PS)")
    for i, row in enumerate(filtered_array):
        print(f"* 行 {i}: {row}")
    
    # 返回筛选后的数组
    return filtered_array

if __name__ == "__main__":
    folder_path = select_folder()
    
    # 选择完文件夹后，弹框让用户输入目标样品编号（支持逗号/空格/分号/中文分隔符/换行）
    def _parse_target_ids(text: str):
        if not text:
            return []
        parts = re.split(r"[，,;；\s]+", text.strip())
        ids = []
        seen = set()
        for p in parts:
            p = p.strip().upper()
            if not p:
                continue
            if p not in seen:
                seen.add(p)
                ids.append(p)
        return ids

    input_root = Tk()
    input_root.withdraw()
    ids_text = simpledialog.askstring(
        title="输入样品编号",
        prompt=(
            "请输入要提取的样品编号：\n"
            "- 多个编号可用逗号/空格/分号/中文逗号/换行分隔\n"
            "- 例如：DX2509530201, DX2509530301"
        ),
        initialvalue="DX2509530201, DX2509530301",
        parent=input_root,
    )
    input_root.destroy()

    target_sample_ids = _parse_target_ids(ids_text or "")
    if not target_sample_ids:
        messagebox.showwarning("未输入编号", "未输入任何样品编号，程序将退出。")
        raise SystemExit(0)

    excel_files = get_excel_files_from_folder(folder_path)
    print("Found Excel files:")
    
    # 创建一个总的数组来存储所有项目的数据
    all_data = []
    
    for excel_file in excel_files:
        print(excel_file)
        result_array = extract_sample_and_concentration(
            excel_file,
            skip_empty_rows=True,
            targets=target_sample_ids,
        )
        
        # 获取文件名并提取汉字部分
        file_name = os.path.basename(excel_file)
        # 使用正则表达式提取汉字
        chinese_chars = re.findall(r'[\u4e00-\u9fff]+', file_name)
        chinese_name = ''.join(chinese_chars) if chinese_chars else file_name
        
        # 基于result_array,再新建一列（第一列），内容为chinese_name
        for row in result_array:
            row.insert(0, chinese_name)
            all_data.append(row)
    
    # 将所有数据写入到一个Excel文件中
    if all_data:
        # 创建DataFrame
        df = pd.DataFrame(all_data, columns=["项目名", "样品编号", "浓度"])
        
        # 弹出文件保存对话框（改为：先选路径，再输入不带后缀的文件名）
        root = Tk()
        root.withdraw()  # 隐藏主窗口

        # 默认文件名（包含关键词以便区分）
        default_filename = "汇总_0953_指定样品提取.xlsx"
        default_basename = os.path.splitext(default_filename)[0]

        # 1) 选择保存文件夹
        messagebox.showinfo(
            "保存提取结果",
            (
                f"即将保存提取结果：\n\n"
                f"数据行数: {len(all_data)} 行\n"
                f"来源文件: {len(excel_files)} 个\n\n"
                f"请先选择保存文件夹，然后输入文件名（不含后缀）。"
            ),
        )
        output_dir = askdirectory(title="选择保存文件夹", initialdir=folder_path)
        if not output_dir:
            print("\n⚠ 用户取消了保存操作（未选择保存文件夹）")
            messagebox.showwarning("取消保存", "未选择保存文件夹，未保存提取结果")
            raise SystemExit(0)

        # 2) 输入不带后缀的文件名
        filename_input = simpledialog.askstring(
            title="输入文件名",
            prompt="请输入要保存的文件名（不含后缀 .xlsx）：",
            initialvalue=default_basename,
            parent=root,
        )

        base_name = (filename_input or "").strip()
        if not base_name:
            messagebox.showinfo("使用默认文件名", f"未输入文件名，将使用默认文件名：{default_filename}")
            base_name = default_basename

        # 3) 组装最终路径并覆盖确认
        output_file = os.path.join(output_dir, f"{base_name}.xlsx")
        if os.path.exists(output_file):
            from tkinter import messagebox as _mb
            if not _mb.askyesno("文件已存在", f"{output_file}\n已存在，是否覆盖？"):
                print("\n⚠ 用户取消了保存操作（拒绝覆盖已存在文件）")
                messagebox.showwarning("取消保存", "未保存提取结果")
                raise SystemExit(0)

        # 4) 保存
        try:
            df.to_excel(output_file, index=False)
            print(f"\n{'='*80}")
            print(f"✅ 所有提取结果已汇总保存到: {output_file}")
            print(f"   共提取 {len(all_data)} 行数据，来自 {len(excel_files)} 个文件")
            print(f"{'='*80}")
            
            # 显示成功提示
            messagebox.showinfo("保存成功", f"提取结果已成功保存到:\n{output_file}")
        except Exception as e:
            print(f"\n❌ 保存失败: {e}")
            messagebox.showerror("保存失败", f"保存失败:\n{e}")
    else:
        print("\n⚠ 没有提取到任何数据")