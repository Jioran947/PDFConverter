# 批量转 PDF 工具

这是一个本地运行的小工具，可以批量把图片、文本、Word、PPT、Excel 等文件转成 PDF。图片和文本转换不需要联网；Office 文件会优先调用你电脑里的 LibreOffice，如果没有 LibreOffice，再尝试调用 Microsoft Office。

## 支持格式

- 图片：JPG、PNG、BMP、GIF、TIFF、WEBP
- 文本：TXT、MD、LOG
- Office：DOC、DOCX、RTF、ODT、PPT、PPTX、ODP、XLS、XLSX、ODS、CSV
- PDF：会复制到输出目录，便于统一整理

## 最简单的使用方式

直接双击：

```text
PDFConverter.exe
```

这个 exe 已经内置 Python 运行环境和必要依赖，不需要你单独安装 Python。

## 新版功能

- 界面已加入 Windows DPI 适配，高清屏下会比上一版清晰。
- 支持中英文切换，默认中文；切换英文后会记住，下次打开仍是英文。
- 支持把文件或文件夹直接拖进文件列表区域。
- 可以先选择转换类型：全部支持格式、Word、PPT、Excel/CSV、图片、文本、已有 PDF。
- 文件列表支持上移、下移、删除、清空。
- 文件列表支持用鼠标直接拖动调整顺序。
- 提供两个转换按钮：一个按文件分别输出 PDF，另一个会在转换后按当前顺序合并为一个 PDF。
- 拖入文件夹时会自动展开成文件，方便调整顺序。
- 图片转 PDF 会按 A4 宽度生成页面，高度按原图比例计算，不再把图片放进 A4 白底中间。
- PPT 转 PDF 后会统一为 A4 宽度，并保持幻灯片原比例。
- 支持拆分 PDF：选中 PDF 后点击“拆分 PDF”，输入页码范围即可导出新 PDF。
- 拆分 PDF 已加入页面预览，可以直接勾选页面后导出。
- 拆分页码输入框和缩略图选择会互相同步：输入页码会选中对应页面，勾选页面也会更新页码范围。
- 拆分预览现在在主窗口内切换显示，提供返回键；预览页左上角叠加选择框，缩略图按上边缘和左边缘对齐。
- 鼠标指针在拆分预览区域内即可用滚轮滚动，不需要移动到滚动条上。
- 拆分预览会根据窗口宽度和页数自动调整缩略图列数与大小；内容不足一屏时滚轮不会让页面上下乱动。
- 拆分预览支持按住 Ctrl 滚动鼠标滚轮或使用触控板调整缩略图大小。
- 拆分预览按每页真实比例显示缩略图，横版页面显示为横版，不会再被放进竖版 A4 白框。
- 调整拆分预览窗口大小时，每行按左上角对齐并重新计算列数和缩略图大小，避免缩略图、边框和页码重叠。
- 拆分预览的每一行会按当前行最高页面决定行高；页码会居中贴近对应缩略图，并与下一行保持清楚间距。
- 批量转换采用临时文件夹中转；只有所有文件都成功后才会输出到目标文件夹，中途失败或取消会删除本次产生的文件。
- 转换、合并、拆分等操作成功输出后，会自动打开输出所在文件夹；如果只输出单个文件，也会额外打开该文件。
- 点击窗口右上角关闭时，会优雅退出程序；若正在转换，会先等当前转换清理完 Office/WPS/LibreOffice 后再关闭，避免残留本程序打开的进程。
- Excel / CSV 会按表格形式转 PDF，并只输出实际有数据的区域。
- 添加不支持的文件或当前类型不匹配的文件时，会弹窗和日志提示，不再静默失败。
- 文件列表为空时，状态栏会显示拖拽提示；关闭过程中的状态提示会跟随当前语言显示。
- 启动窗口会按屏幕大小自动放大；缩小窗口时会优先压缩下方日志区，尽量保留文件拖拽区域。
- Windows 下调用 Microsoft Office / WPS 转换时，会在后台线程内初始化并释放 COM，减少 `尚未调用 CoInitialize` 和残留 Office 进程问题。

## 开发版使用

1. 双击 `install.bat` 安装依赖。
2. 双击 `start.bat` 打开窗口版。
3. 添加文件或文件夹，选择输出位置，然后点“开始转换”。

如果你的电脑没有 Python，请先安装 Python 3.10 或更新版本：

https://www.python.org/downloads/windows/

## 重新打包成 exe

如果你修改了 `pdf_converter.py`，可以双击：

```text
build_app.bat
```

打包后的文件会生成在：

```text
dist\PDFConverter.exe
```

## Office 文件转换

如果要转换 Word、PPT、Excel，推荐安装免费的 LibreOffice：

https://www.libreoffice.org/download/download-libreoffice/

安装后程序会自动寻找 LibreOffice 并调用它转换。你也可以使用 Microsoft Office 或 WPS；当前 exe 已经内置了调用 Office/WPS 所需的 Python 组件，但电脑上仍然需要安装 Word、PowerPoint、Excel 或 WPS 本体，并且 WPS 需要正常注册 COM 自动化组件。

转换优先级：

1. LibreOffice
2. Microsoft Office COM
3. WPS COM

## 命令行批量转换

转换一个文件夹：

```powershell
python pdf_converter.py "D:\待转换文件" -o "D:\PDF输出"
```

把多张图片合并为一个 PDF：

```powershell
python pdf_converter.py "D:\图片文件夹" -o "D:\PDF输出" --merge-images --merged-name "相册.pdf"
```

只处理当前文件夹，不包含子文件夹：

```powershell
python pdf_converter.py "D:\待转换文件" -o "D:\PDF输出" --no-recursive
```

## 输出位置

默认输出到系统“下载”文件夹。如果输出文件名重复，程序会自动加 `_2`、`_3`，避免覆盖原文件。
