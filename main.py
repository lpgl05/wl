import os
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
from datetime import datetime
from tkinter import Tk, messagebox, simpledialog
from tkinter.filedialog import askdirectory, asksaveasfilename, askopenfilename

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
    - 去除"平行X"字样
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

def _parse_target_ids(text: str):
    """
    将文本字符串解析为编号列表。
    支持逗号/空格/分号/中文逗号/换行等分隔符，自动去重并转大写。
    
    参数:
        text: 多个编号的文本（以分隔符分开）
    
    返回:
        去重后的编号列表
    """
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

def read_ids_from_txt(file_path: str):
    """
    从TXT文件读取编号列表，每行一个编号。
    自动过滤空行和注释行，去重并转大写。
    
    参数:
        file_path: TXT文件的完整路径
    
    返回:
        去重后的编号列表，若读取失败返回 None
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        ids = []
        seen = set()
        for line in lines:
            line = line.strip()
            # 跳过空行和注释行（以#开头）
            if not line or line.startswith('#'):
                continue
            # 去重并转大写
            line_upper = line.upper()
            if line_upper not in seen:
                seen.add(line_upper)
                ids.append(line_upper)
        
        return ids if ids else None
    except Exception as e:
        messagebox.showerror("读取失败", f"无法读取TXT文件：\n{e}")
        return None

def extract_metadata_from_excel(file_path):
    """
    从Excel文件中提取元数据信息（分析人、分析日期、使用仪器、仪器型号）
    通过查找这些标签名称，然后获取其右侧相邻单元格的值
    
    参数:
        file_path: Excel文件的完整路径
    
    返回:
        字典，包含 {
            'analyzer': '分析人',
            'analysis_date': '分析日期',
            'instrument': '使用仪器',
            'instrument_model': '仪器型号'
        }
    """
    metadata = {
        'analyzer': '',
        'analysis_date': ''
    }
    
    try:
        if file_path.endswith('.xlsx'):
            wb = load_workbook(file_path, data_only=True)
            if not wb.sheetnames:
                return metadata
            ws = wb[wb.sheetnames[0]]
            
            # 标签和对应字段的映射
            label_mapping = {
                '分析人': 'analyzer',
                '分析人员': 'analyzer',
                '分析日期': 'analysis_date'
            }
            
            # 遍历所有单元格查找标签
            for row in ws.iter_rows(min_row=1, max_row=100):
                for cell in row:
                    cell_value = str(cell.value).strip() if cell.value else ""
                    if cell_value in label_mapping:
                        # 找到标签，获取右侧相邻单元格的值
                        right_cell = ws.cell(row=cell.row, column=cell.column + 1)
                        field_name = label_mapping[cell_value]
                        metadata[field_name] = str(right_cell.value).strip() if right_cell.value else ""
        else:
            # .xls 文件处理
            df = pd.read_excel(file_path, header=None)
            label_mapping = {
                '分析人': 'analyzer',
                '分析人员': 'analyzer',
                '分析日期': 'analysis_date'
            }
            
            for idx, row in df.iterrows():
                for col_idx, cell_value in enumerate(row):
                    cell_str = str(cell_value).strip() if cell_value and pd.notna(cell_value) else ""
                    if cell_str in label_mapping:
                        # 获取右侧相邻单元格
                        if col_idx + 1 < len(row):
                            right_value = row.iloc[col_idx + 1]
                            field_name = label_mapping[cell_str]
                            metadata[field_name] = str(right_value).strip() if right_value and pd.notna(right_value) else ""
    except Exception as e:
        print(f"⚠ 提取元数据失败: {e}")
    
    return metadata

def ask_sample_ids_source():
    """
    弹出对话框让用户选择样品编号的输入方式：
    1. 直接输入 - 在对话框中手动输入多个编号
    2. 从TXT文件导入 - 选择一个TXT文件，每行一个编号
    
    返回:
        编号列表，若用户取消则返回 None
    """
    root = Tk()
    root.withdraw()
    
    # 先让用户选择输入方式
    choice = messagebox.askyesnocancel(
        title="选择样品编号输入方式",
        message="请选择如何输入样品编号：\n"
                "\n"
                "\"是\" - 直接输入：在对话框中输入多个编号（逗号/空格/分号分隔）\n"
                "\"否\" - 从文件导入：选择一个TXT文件，每行一个编号\n"
                "\"取消\" - 退出程序",
    )
    
    if choice is None:  # 取消
        root.destroy()
        return None
    elif choice:  # 是 - 直接输入
        ids_text = simpledialog.askstring(
            title="输入样品编号",
            prompt=(
                "请输入要提取的样品编号：\n"
                "- 多个编号可用逗号/空格/分号/中文逗号/换行分隔\n"
                "- 例如：DX2509530201, DX2509530301"
            ),
            initialvalue="DX2509530201, DX2509530301",
            parent=root,
        )
        root.destroy()
        if ids_text:
            return _parse_target_ids(ids_text)
        else:
            return None
    else:  # 否 - 从文件导入
        root.destroy()
        file_path = askopenfilename(
            title="选择包含样品编号的TXT文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if file_path:
            return read_ids_from_txt(file_path)
        else:
            return None

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
    concentration_col_keyword = '浓度'  # 只需包含"浓度"字样即可
    
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
            
            # 查找浓度列（只需包含"浓度"字样）
            temp_concentration_col = None
            for col_name in temp_df.columns:
                if concentration_col_keyword in col_name:
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
            error_msg = str(e)
            # 如果是 xlrd 相关错误，说明文件是 .xls 格式但缺少 xlrd 库
            if 'xlrd' in error_msg.lower():
                print(f"❌ 文件格式错误: 该文件是 .xls 格式，需要安装 xlrd 库来支持")
                print(f"   解决方案: pip install xlrd 或 uv add xlrd")
                return []
            # 其他异常继续尝试
            print(f"尝试 header={header} 失败: {e}")
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

    # 让用户选择样品编号输入方式：直接输入或从TXT文件导入
    target_sample_ids = ask_sample_ids_source()
    if not target_sample_ids:
        messagebox.showwarning("未输入编号", "未输入任何样品编号，程序将退出。")
        raise SystemExit(0)

    excel_files = get_excel_files_from_folder(folder_path)
    print("Found Excel files:")
    print(f"待处理的样品编号（共 {len(target_sample_ids)} 个）: {', '.join(target_sample_ids)}")
    
    # 创建一个总的数组来存储所有项目的数据
    all_data = []
    
    for excel_file in excel_files:
        print(excel_file)
        
        # 提取元数据（分析人、分析日期、使用仪器、仪器型号）
        metadata = extract_metadata_from_excel(excel_file)
        
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
        
        # 基于result_array，添加元数据列
        for row in result_array:
            # 按新的顺序插入：分析人、分析日期、项目名、样品编号、浓度
            row.insert(0, metadata.get('analysis_date', ''))
            row.insert(0, metadata.get('analyzer', ''))
            row.insert(2, chinese_name)
            all_data.append(row)
    
    # 将所有数据写入到一个Excel文件中
    if all_data:
        # 创建DataFrame，包含新增的元数据列
        df = pd.DataFrame(all_data, columns=[
            "分析人", 
            "分析日期", 
            "项目名", 
            "样品编号", 
            "浓度"
        ])
        
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