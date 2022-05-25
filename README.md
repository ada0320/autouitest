# autouitest python自动化UI测试工具

# 使用方法:

在excel中写好测试步骤后运行 start.bat

(详见demo)
命令行启动：

python wxUI.py 为带UI的启动
python AutoUITest.py 为控制台启动

# 环境配置：

1, 安装python 3.9.10 (因为目前wxPython暂不支持更高版本), 并配置系统环境变量Path


2, 命令行安装以下python库

pip install xlrd==1.2.0

pip install pillow

pip install wxpython

pip install pyperclip

pip install pyautogui

pip install opencv-python -i https://pypi.tuna.tsinghua.edu.cn/simple


如果某个库安装失败,可以手动下载对应 .whl 文件, 解压缩后, 放入{python安装目录}\Lib\site-packages
