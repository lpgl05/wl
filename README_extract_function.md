# extract_sample_and_concentration 函数使用说明

## 函数签名
```python
extract_sample_and_concentration(file_path, skip_empty_rows=False)
```

## 参数说明
- `file_path` (str): Excel文件的完整路径
- `skip_empty_rows` (bool, 可选): 是否跳过样品编号和浓度都为空的行，默认False

## 返回值
返回一个二维数组 (list)，每行包含 `[样品编号, 浓度]`

## 功能特点
1. **自动检测列名所在行**: 自动尝试第1、2、3行作为列名
2. **灵活匹配列名**:
   - 样品编号列: 匹配 `样品编号`、`样品` 或 `编号`
   - 浓度列: 匹配 `浓度mg/L`、`浓度mg/l` 或 `浓度`
3. **可选的空行过滤**: 可以过滤掉两列都为空的行

## 使用示例

### 示例1: 基本用法（包含所有行）
```python
from main import extract_sample_and_concentration

file_path = r"D:\data\your_file.xlsx"
result = extract_sample_and_concentration(file_path)

print(f"提取了 {len(result)} 行数据")
for row in result:
    print(row)  # [样品编号, 浓度]
```

### 示例2: 过滤空行
```python
from main import extract_sample_and_concentration

file_path = r"D:\data\your_file.xlsx"
result = extract_sample_and_concentration(file_path, skip_empty_rows=True)

print(f"提取了 {len(result)} 行有效数据")
for i, row in enumerate(result):
    sample_no, concentration = row
    print(f"{i+1}. 样品编号: {sample_no}, 浓度: {concentration}")
```

### 示例3: 处理返回的数据
```python
from main import extract_sample_and_concentration
import pandas as pd

file_path = r"D:\data\your_file.xlsx"
result = extract_sample_and_concentration(file_path, skip_empty_rows=True)

# 转换为pandas DataFrame进行进一步处理
df = pd.DataFrame(result, columns=['样品编号', '浓度'])
print(df)

# 过滤掉浓度为nan的行
df_clean = df.dropna(subset=['浓度'])
print(f"\n有浓度值的数据: {len(df_clean)} 行")
```

## 实际测试结果

### 测试文件: 0953-DX-挥发酚.xlsx

**不过滤空行（默认）**:
- 提取 74 行数据（包含空行）

**过滤空行**:
- 提取 18 行有效数据

### 输出示例
```
✓ 使用第 3 行作为列名 (header=2)
✓ 找到样品编号列: '编号'
✓ 找到浓度列: '浓度mg/L'

提取的二维数组 (共 18 行):
列顺序: [样品编号, 浓度]

数据预览 (前10行):
  行 0: ['空白1', nan]
  行 1: ['空白2', nan]
  行 2: ['DX2509530101平行1', '<0.0003']
  行 3: ['DX2509530101平行2', nan]
  行 4: ['DX2509530201', '<0.0003']
  ...
```

## 注意事项
1. 函数会自动在控制台打印处理进度和结果预览
2. 如果找不到目标列，会返回空数组 `[]`
3. 数据中的 `nan` 表示Excel中的空单元格
4. 浓度值可能是数字或字符串（如 `'<0.0003'`）
