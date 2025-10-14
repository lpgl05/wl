import re
import os

# 测试文件名提取汉字
test_files = [
    "0953-DX-总硬度.xlsx",
    "0953-DX-挥发酚.xlsx",
    "0953-DX-氰化物.xlsx",
    "0953-DX氨氮.xlsx"
]

print("测试提取文件名中的汉字:")
print("="*60)
for file_name in test_files:
    # 使用正则表达式提取汉字
    chinese_chars = re.findall(r'[\u4e00-\u9fff]+', file_name)
    chinese_name = ''.join(chinese_chars) if chinese_chars else file_name
    print(f"{file_name:30} -> {chinese_name}")
