@echo off
chcp 65001 > nul
echo.
echo ====================================================
echo         Excel到Word模板转换工具 - 自动化打包
echo ====================================================
echo.

echo 正在检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python环境，请先安装Python 3.7以上版本
    pause
    exit /b 1
)

echo 正在检查pip...
pip --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到pip，请检查Python安装
    pause
    exit /b 1
)

echo 正在安装/升级依赖包...
pip install pandas python-docx openpyxl xlrd pillow lxml pyinstaller

echo.
echo 正在清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del *.spec
if exist __pycache__ rmdir /s /q __pycache__

echo.
echo 正在执行打包...
pyinstaller --onefile --windowed --clean --noconfirm ^
    --name "Excel到Word模板转换工具" ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module numpy.testing ^
    --exclude-module pytest ^
    --exclude-module IPython ^
    --exclude-module jupyter ^
    --hidden-import pandas ^
    --hidden-import docx ^
    --hidden-import openpyxl ^
    --hidden-import xlrd ^
    --hidden-import PIL ^
    --hidden-import lxml ^
    --hidden-import tkinter ^
    --hidden-import tkinter.ttk ^
    --hidden-import tkinter.filedialog ^
    --hidden-import tkinter.messagebox ^
    excel2word_template_version_1.py

if errorlevel 1 (
    echo.
    echo 打包失败，请检查错误信息
    pause
    exit /b 1
)

echo.
echo 正在创建使用说明...
(
echo # Excel到Word模板转换工具
echo.
echo 作者：yf
echo 年份：2025
echo 开源协议：MIT License
echo.
echo ## 使用说明
echo.
echo 1. 直接双击 "Excel到Word模板转换工具.exe" 运行
echo 2. 无需安装Python环境或其他依赖
echo 3. 首次运行可能需要较长时间加载
echo.
echo ## 系统要求
echo - Windows 7 及以上版本
echo - 约100MB磁盘空间
echo - 2GB以上内存（推荐）
echo.
echo ## 程序功能
echo - 导入Excel数据文件
echo - 导入Word模板文件
echo - 自动识别模板中的占位符
echo - 字段映射和批量替换
echo - 图片插入功能
echo - 保持原始格式和样式
echo - 批量生成Word文档
echo.
echo ## 支持的文件格式
echo - Excel: .xlsx, .xls
echo - Word: .docx
echo - 图片: .jpg, .jpeg, .png, .bmp, .gif
echo.
echo ## 开源协议
echo 本软件遵循MIT开源协议：
echo - 可以自由使用、修改、分发
echo - 可以用于商业项目
echo - 需要保留原作者版权声明
echo - 作者不承担任何责任和担保
echo.
echo ## 注意事项
echo - 杀毒软件可能误报，请添加信任
echo - 处理大量数据时请耐心等待
echo.
echo 如有问题，请查看程序内的"使用助手"功能。
) > dist\README.txt

echo.
echo ====================================================
echo                    打包完成！
echo ====================================================
echo.
echo 输出文件位置：
echo - 可执行文件: dist\Excel到Word模板转换工具.exe
echo - 使用说明: dist\README.txt
echo.
echo 建议：
echo 1. 在其他电脑上测试exe文件
echo 2. 将整个dist文件夹打包发给用户
echo 3. 提醒用户阅读README.txt
echo.
echo 正在打开输出目录...
start dist

pause 