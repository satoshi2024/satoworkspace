你是 Windows 11，直接用 CMD + winget 安装最方便。Microsoft 当前给出的命令行安装方式是安装 Python 3.14；Python 官网也将 3.14 列为当前稳定版本。

一、打开 CMD

按：

Win + R

输入：

cmd

按回车。

不需要先用管理员模式；安装时出现权限窗口再确认即可。

二、确认 winget 能使用

在 CMD 输入：

winget --version

显示类似下面的版本号即可：

v1.x.x

WinGet 是 Windows 10、Windows 11 的命令行软件包管理工具。

三、安装 Python

复制下面整条命令到 CMD：

winget install --id Python.Python.3.14 -e --source winget --accept-package-agreements --accept-source-agreements

等待出现安装成功的信息。

安装完成后，关闭当前 CMD 窗口，再重新打开一个 CMD。这一步很重要，否则环境变量可能还没有刷新。

四、确认 Python 安装成功

重新打开 CMD 后依次执行：

py --version
python --version

正常情况下会显示类似：

Python 3.14.x

继续确认 pip：

py -m pip --version

五、升级 pip

py -m pip install --upgrade pip

六、安装程序需要的 pywin32

我给你的界面程序使用 Excel 的 Windows COM 接口，所以需要安装 pywin32：

py -m pip install pywin32

安装完成后测试：

py -c "import win32com.client; print('pywin32 安装成功')"

正常会显示：

pywin32 安装成功

tkinter 是 Windows 官方 Python 安装中自带的界面组件，不需要另外安装。

七、建立程序文件夹

在 CMD 中执行：

mkdir C:\ExcelTextColorTool

进入文件夹：

cd /d C:\ExcelTextColorTool

八、创建 Python 源码文件

执行：

notepad excel_text_color_gui.py

Windows 会提示文件不存在，询问是否创建，选择：

是

把我上一条消息提供的全部 Python 源码复制进去，然后在记事本中按：

Ctrl + S

保存并关闭记事本。

确认文件存在：

dir

应该可以看到：

excel_text_color_gui.py

注意不能变成：

excel_text_color_gui.py.txt

九、运行界面程序

在 CMD 中执行：

cd /d C:\ExcelTextColorTool
py excel_text_color_gui.py

这时会打开“Excel 指定文字变色工具”的窗口。

使用顺序是：

1. 选择 Excel 文件
2. 选择工作表
3. 输入单元格，例如 BG3
4. 或输入范围，例如 BG3:BG24
5. 输入指定文字
6. 选择颜色
7. 点击“开始处理”

十、以后快速启动

每次运行只需要打开 CMD，输入：

cd /d C:\ExcelTextColorTool
py excel_text_color_gui.py

也可以创建一个双击启动文件。在 CMD 执行：

cd /d C:\ExcelTextColorTool
echo @echo off>启动工具.bat
echo cd /d "%%~dp0">>启动工具.bat
echo py excel_text_color_gui.py>>启动工具.bat
echo pause>>启动工具.bat

以后打开：

C:\ExcelTextColorTool

双击：

启动工具.bat

即可运行。

常见错误

winget 不是内部或外部命令

先运行 Windows 更新，并在 Microsoft Store 中更新“应用安装程序”，然后重新打开 CMD。

也可以使用 Python 官方的安装管理器命令：

winget install 9NQ7512CXL7T

这是 Python 官方提供的 Windows Python Install Manager 安装方式。

py 不是内部或外部命令，但 python 可以使用

把后面的命令从：

py -m pip install pywin32

改成：

python -m pip install pywin32

运行程序也改成：

python excel_text_color_gui.py

输入 python 后跳到 Microsoft Store

优先使用：

py --version

程序也使用：

py excel_text_color_gui.py

提示 No module named win32com

重新执行：

py -m pip install --upgrade pywin32

然后关闭 CMD，重新打开再运行。

程序打开 Excel 失败

确认电脑中已经安装桌面版 Microsoft Excel，并关闭正在编辑的目标 Excel 文件。