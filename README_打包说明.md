# Excel数据提取工具 - 打包说明

## 📦 项目打包指南

本项目使用 PyInstaller 将Python程序打包成Windows可执行文件(.exe)，可在Windows 7/10/11上直接运行，无需安装Python环境。

---

## 🚀 快速打包

### 方法1: 使用自动化脚本（推荐）

在项目根目录运行：

```bash
# 使用uv项目的Python环境
D:/code/vscode/wl/.venv/Scripts/python.exe build.py
```

脚本会自动完成：
1. ✅ 安装 PyInstaller
2. ✅ 打包程序
3. ✅ 显示结果信息

### 方法2: 手动打包

```bash
# 1. 安装PyInstaller
python -m pip install pyinstaller

# 2. 打包
pyinstaller build.spec --clean
```

---

## 📁 打包后的文件结构

```
wl/
├── dist/
│   └── Excel数据提取工具.exe  ← 最终可执行文件
├── build/                      ← 临时构建文件
├── main.py                     ← 源代码
├── build.py                    ← 打包脚本
└── build.spec                  ← PyInstaller配置
```

---

## 💡 使用可执行文件

### 在开发机上测试

```bash
# 运行生成的exe
.\dist\Excel数据提取工具.exe
```

### 分发给用户

1. 将 `dist/Excel数据提取工具.exe` 复制到目标电脑
2. 双击运行
3. 选择包含Excel文件的文件夹
4. 程序自动处理并提示保存结果

---

## ⚙️ 配置说明

### build.spec 关键配置

| 配置项 | 说明 | 当前值 |
|--------|------|--------|
| `name` | 生成的exe文件名 | `Excel数据提取工具` |
| `console` | 是否显示控制台 | `True` (显示处理进度) |
| `upx` | 是否压缩 | `True` (减小文件大小) |
| `icon` | 应用图标 | `None` (可自定义) |

### 添加自定义图标

1. 准备一个 `.ico` 格式的图标文件（建议256x256）
2. 放在项目根目录，例如 `icon.ico`
3. 修改 `build.spec` 中的 `icon` 参数：
   ```python
   icon='icon.ico'
   ```

---

## 🎯 兼容性

### 支持的操作系统

- ✅ Windows 7 (SP1及以上)
- ✅ Windows 10
- ✅ Windows 11

### 系统要求

- 内存: 至少 2GB RAM
- 磁盘: 至少 100MB 可用空间
- .NET Framework 4.0+ (Windows 7需要)

---

## 🔧 常见问题

### Q1: 打包后文件太大？

**解决方案：**
- 已启用UPX压缩
- 可以排除不需要的依赖
- 在 `build.spec` 中添加到 `excludes` 列表

### Q2: 杀毒软件报毒？

**原因：** PyInstaller打包的文件可能被误报

**解决方案：**
- 添加到杀毒软件白名单
- 使用代码签名证书签名exe文件
- 从官方渠道分发

### Q3: 双击exe无反应？

**排查步骤：**
1. 检查是否有杀毒软件拦截
2. 以管理员身份运行
3. 查看是否有错误日志
4. 将 `console=True` 以查看错误信息

### Q4: 在其他电脑上运行报错？

**可能原因：**
- 缺少 Visual C++ 运行库
- 系统版本太低

**解决方案：**
- 安装 [Microsoft Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe)
- 确保系统已更新

---

## 📝 打包优化建议

### 减小文件大小

1. **排除不必要的模块**
   ```python
   excludes=[
       'matplotlib',
       'scipy',
       'numpy.testing',
   ]
   ```

2. **单文件模式** (可选)
   将所有文件打包成一个exe：
   ```python
   exe = EXE(
       ...,
       onefile=True,  # 添加这一行
   )
   ```

### 加快启动速度

- 使用 `--noupx` 选项（牺牲文件大小）
- 使用文件夹模式而非单文件模式

---

## 🛠️ 高级配置

### 隐藏控制台窗口

适合不需要看到处理过程的场景：

```python
exe = EXE(
    ...,
    console=False,  # 修改为False
)
```

### 添加版本信息

创建 `version_info.txt`：
```
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(1, 0, 0, 0),
    prodvers=(1, 0, 0, 0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        u'040904B0',
        [StringStruct(u'CompanyName', u'Your Company'),
        StringStruct(u'FileDescription', u'Excel数据提取工具'),
        StringStruct(u'FileVersion', u'1.0.0.0'),
        StringStruct(u'ProductName', u'Excel数据提取工具'),
        StringStruct(u'ProductVersion', u'1.0.0.0')])
    ]),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
```

然后在 `build.spec` 中添加：
```python
exe = EXE(
    ...,
    version='version_info.txt',
)
```

---

## 📊 打包性能对比

| 模式 | 文件数量 | 文件大小 | 启动速度 |
|------|---------|---------|---------|
| 单文件 | 1个 | ~30MB | 较慢 |
| 文件夹 | 多个 | ~40MB | 较快 |

**当前配置**: 单文件模式（便于分发）

---

## 🔐 安全建议

1. **代码签名**: 为exe文件签名以提高信任度
2. **病毒扫描**: 打包后进行病毒扫描
3. **完整性验证**: 提供MD5/SHA256校验值
4. **官方渠道**: 通过官方渠道分发

---

## 📮 技术支持

如有问题，请检查：
1. Python版本是否 >= 3.13
2. 所有依赖是否正确安装
3. PyInstaller版本是否最新
4. Windows系统是否已更新

---

## 📄 许可证

本打包配置基于项目许可证。

---

**最后更新**: 2025-10-15
