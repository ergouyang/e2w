#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel到Word模板转换工具 - 自动化打包脚本

Author: yf
Year: 2025
License: MIT License
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def check_dependencies():
    """检查依赖是否安装"""
    print("检查依赖...")
    
    required_packages = [
        'pandas',
        'python-docx', 
        'openpyxl',
        'xlrd',
        'pillow',
        'lxml',
        'pyinstaller'
    ]
    
    missing_packages = []
    for package in required_packages:
        try:
            __import__(package.replace('-', '_'))
            print(f"✓ {package} 已安装")
        except ImportError:
            missing_packages.append(package)
            print(f"✗ {package} 未安装")
    
    if missing_packages:
        print(f"\n缺少以下依赖包：{missing_packages}")
        print("请运行以下命令安装：")
        print(f"pip install {' '.join(missing_packages)}")
        return False
    
    print("所有依赖检查通过！")
    return True

def clean_build_dirs():
    """清理构建目录"""
    print("清理构建目录...")
    
    dirs_to_clean = ['build', 'dist', '__pycache__']
    files_to_clean = ['*.spec']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"删除目录: {dir_name}")
    
    # 清理spec文件
    for spec_file in Path('.').glob('*.spec'):
        spec_file.unlink()
        print(f"删除文件: {spec_file}")

def create_pyinstaller_spec():
    """创建PyInstaller配置文件"""
    print("创建PyInstaller配置...")
    
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['excel2word_template_version_1.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'pandas',
        'docx',
        'openpyxl',
        'xlrd',
        'PIL',
        'lxml',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'difflib',
        'glob',
        'tempfile',
        'subprocess',
        'platform',
        'datetime',
        'copy',
        'typing',
        'pandas.plotting',
        'pandas.io.formats.excel',
        'openpyxl.cell.cell',
        'openpyxl.styles',
        'docx.shared',
        'docx.enum.text',
        'docx.oxml.ns',
        'xml.etree.ElementTree',
        'PIL.Image',
        'PIL.ImageTk'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'numpy.testing',
        'pytest',
        'IPython',
        'jupyter'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Excel到Word模板转换工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
    version_info=None
)
'''
    
    with open('excel2word_converter.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("配置文件创建完成: excel2word_converter.spec")

def build_exe():
    """构建exe文件"""
    print("开始构建exe文件...")
    
    # 使用spec文件构建
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--clean',
        '--noconfirm',
        'excel2word_converter.spec'
    ]
    
    print(f"执行命令: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("构建成功！")
        print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"构建失败: {e}")
        print(f"错误输出: {e.stderr}")
        return False

def optimize_exe():
    """优化exe文件"""
    print("优化exe文件...")
    
    exe_path = Path('dist/Excel到Word模板转换工具.exe')
    if exe_path.exists():
        file_size = exe_path.stat().st_size / (1024 * 1024)  # MB
        print(f"exe文件大小: {file_size:.2f} MB")
        
        if file_size > 100:
            print("警告: exe文件较大，可能需要优化")
            print("建议:")
            print("1. 检查是否包含了不必要的库")
            print("2. 使用--exclude-module排除不需要的模块")
            print("3. 考虑使用onedir模式而非onefile模式")
    
    return True

def create_readme():
    """创建使用说明"""
    print("创建使用说明...")
    
    readme_content = '''# Excel到Word模板转换工具

作者：yf
年份：2025
开源协议：MIT License

## 使用说明

1. **运行程序**
   - 直接双击 `Excel到Word模板转换工具.exe` 运行
   - 无需安装Python环境或其他依赖

2. **系统要求**
   - Windows 7 及以上版本
   - 约100MB磁盘空间
   - 2GB以上内存（推荐）

3. **程序功能**
   - 导入Excel数据文件
   - 导入Word模板文件
   - 自动识别模板中的占位符
   - 字段映射和批量替换
   - 图片插入功能
   - 保持原始格式和样式
   - 批量生成Word文档

4. **支持的文件格式**
   - Excel: .xlsx, .xls
   - Word: .docx
   - 图片: .jpg, .jpeg, .png, .bmp, .gif

5. **使用步骤**
   - 第一步：选择Excel文件和Word模板
   - 第二步：配置字段映射
   - 第三步：设置图片映射（可选）
   - 第四步：选择导出设置
   - 第五步：预览或直接导出

6. **注意事项**
   - 首次运行可能需要较长时间加载
   - 杀毒软件可能报告误报，请添加信任
   - 处理大量数据时请耐心等待

## 开源协议

本软件遵循MIT开源协议：
- 可以自由使用、修改、分发
- 可以用于商业项目
- 需要保留原作者版权声明
- 作者不承担任何责任和担保

## 技术支持

如有问题，请查看程序内的"使用助手"或联系开发者。

---
构建时间: {build_time}
'''
    
    import datetime
    build_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    with open('dist/README.txt', 'w', encoding='utf-8') as f:
        f.write(readme_content.format(build_time=build_time))
    
    print("使用说明创建完成: dist/README.txt")

def main():
    """主函数"""
    print("=" * 50)
    print("Excel到Word模板转换工具 - 自动化打包")
    print("=" * 50)
    
    # 检查依赖
    if not check_dependencies():
        return False
    
    # 清理构建目录
    clean_build_dirs()
    
    # 创建配置文件
    create_pyinstaller_spec()
    
    # 构建exe
    if not build_exe():
        return False
    
    # 优化exe
    optimize_exe()
    
    # 创建使用说明
    create_readme()
    
    print("\n" + "=" * 50)
    print("打包完成！")
    print("=" * 50)
    print("输出文件:")
    print("- 可执行文件: dist/Excel到Word模板转换工具.exe")
    print("- 使用说明: dist/README.txt")
    print("\n建议:")
    print("1. 在其他电脑上测试exe文件")
    print("2. 将整个dist文件夹打包发给用户")
    print("3. 提醒用户阅读README.txt")
    
    return True

if __name__ == "__main__":
    try:
        success = main()
        if success:
            print("\n按回车键退出...")
            input()
        else:
            print("\n打包失败，按回车键退出...")
            input()
    except KeyboardInterrupt:
        print("\n用户中断操作")
    except Exception as e:
        print(f"\n打包过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        print("\n按回车键退出...")
        input() 