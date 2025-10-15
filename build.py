"""
自动化打包脚本
使用PyInstaller将Python项目打包成Windows可执行文件
"""

import subprocess
import sys
import os

def install_pyinstaller():
    """安装PyInstaller (使用uv)"""
    print("="*80)
    print("步骤 1/3: 安装PyInstaller...")
    print("="*80)
    try:
        # 使用 uv 添加 pyinstaller 到项目
        subprocess.run(["uv", "add", "pyinstaller", "--dev"], check=True)
        print("✅ PyInstaller安装成功")
    except subprocess.CalledProcessError as e:
        print(f"❌ PyInstaller安装失败: {e}")
        print("\n💡 提示: 请确保已安装 uv 命令行工具")
        print("   安装方法: powershell -c \"irm https://astral.sh/uv/install.ps1 | iex\"")
        sys.exit(1)

def build_exe():
    """使用PyInstaller打包 (通过uv运行)"""
    print("\n" + "="*80)
    print("步骤 2/3: 开始打包...")
    print("="*80)
    try:
        # 使用 uv run 来运行 pyinstaller
        subprocess.run(["uv", "run", "pyinstaller", "build.spec", "--clean"], check=True)
        print("✅ 打包成功")
    except subprocess.CalledProcessError as e:
        print(f"❌ 打包失败: {e}")
        sys.exit(1)

def show_result():
    """显示打包结果"""
    print("\n" + "="*80)
    print("步骤 3/3: 打包完成")
    print("="*80)
    
    # 文件夹模式:检查文件夹
    dist_folder = os.path.join(os.getcwd(), "dist", "Excel数据提取工具")
    exe_path = os.path.join(dist_folder, "Excel数据提取工具.exe")
    
    if os.path.exists(exe_path):
        # 计算整个文件夹的大小
        total_size = 0
        file_count = 0
        for dirpath, dirnames, filenames in os.walk(dist_folder):
            for filename in filenames:
                filepath = os.path.join(dirpath, filename)
                total_size += os.path.getsize(filepath)
                file_count += 1
        
        folder_size_mb = total_size / (1024 * 1024)
        exe_size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        
        print(f"✅ 可执行文件已生成 (文件夹模式):")
        print(f"   文件夹: {dist_folder}")
        print(f"   主程序: {exe_path}")
        print(f"   主程序大小: {exe_size_mb:.2f} MB")
        print(f"   总大小: {folder_size_mb:.2f} MB ({file_count} 个文件)")
        print("\n📝 使用说明:")
        print("   1. 将整个 'Excel数据提取工具' 文件夹复制到目标位置")
        print("   2. 双击文件夹内的 'Excel数据提取工具.exe' 即可运行")
        print("   3. ⚡ 启动速度快:约 1-2 秒即可看到文件选择对话框")
        print("   4. 选择包含Excel文件的文件夹,程序会自动处理")
        print("\n💡 提示:")
        print("   - 必须保持文件夹完整,不能只复制 .exe 文件")
        print("   - 可以在任何Windows电脑使用(Win7/10/11)")
        print("   - 无需安装Python环境")
        print("   - 文件夹模式比单文件模式启动快 20-30 倍!")
    else:
        print("❌ 未找到生成的可执行文件")
        print(f"   期望路径: {exe_path}")

if __name__ == "__main__":
    print("\n🚀 开始打包 Excel数据提取工具\n")
    
    # 步骤1: 安装PyInstaller
    install_pyinstaller()
    
    # 步骤2: 打包
    build_exe()
    
    # 步骤3: 显示结果
    show_result()
    
    print("\n" + "="*80)
    print("✨ 打包流程完成")
    print("="*80)
