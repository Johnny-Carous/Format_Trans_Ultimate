# 文档格式转换工具

一个简单易用的图形界面工具，用于将 **PDF/EPUB** 文档转换为 **TXT、MD、DOCX、PDF、EPUB** 等格式。内置 OCR 识别（自动处理扫描件），支持手动截图 PDF，转换后自动清空界面，方便连续使用。

## ✨ 功能特点

- **多格式转换**：支持 PDF、EPUB 转换为 `txt`、`md`、`docx`、`pdf`、`epub` 等格式（docx/pdf/epub 需 Pandoc 支持）。
- **自动 OCR**：扫描件 PDF 会自动调用 Tesseract 进行文字识别，用户无感知。
- **手动截图**：在 PDF 页面上框选任意区域保存为图片（独立功能）。
- **连续使用**：转换成功后自动清空源文件路径，焦点回到“浏览”按钮，立即选择下一个文件。
- **纯净输出**：输出文件仅包含纯文本内容，无页码标记或额外格式。
- **依赖自动安装**：首次运行会自动安装所需的 Python 包。

## 📦 安装要求

### 1. Python 环境
- 需要 Python 3.7 或更高版本。
- 建议使用虚拟环境（可选）。

### 2. 外部工具（可选但推荐）
| 工具 | 作用 | 下载/安装方式 |
|------|------|--------------|
| **Tesseract-OCR** | 提供 OCR 识别能力 | [UB Mannheim 下载页](https://github.com/UB-Mannheim/tesseract/wiki)（安装时勾选中文语言包） |
| **Poppler** | 提升 OCR 性能（pdf2image 用） | [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases)（解压后添加到 PATH） |
| **Pandoc** | 支持 docx/pdf/epub 等格式转换 | 程序会自动安装 `pypandoc_binary`（无需手动安装） |

> **注意**：如果未安装 Tesseract，OCR 功能将不可用，但仍可转换文字型 PDF/EPUB。如果未安装 Pandoc，程序会自动使用内置的 `pypandoc_binary`，无需用户干预。

## 🚀 使用方法

1. **运行程序**  
   双击脚本文件（或 `python 脚本名.py`）启动图形界面。

2. **选择源文件**  
   点击“浏览”按钮，选择 PDF 或 EPUB 文件。

3. **选择目标格式**  
   从下拉列表中选择要转换的格式（txt、md、docx、pdf、epub）。  
   > docx/pdf/epub 需要 Pandoc 支持，程序会自动处理。

4. **选择输出目录**  
   默认输出到当前目录下的 `output` 文件夹，可点击“浏览”自定义。

5. **开始转换**  
   点击“开始转换”，等待进度条完成。转换成功后：
   - 弹出提示框显示保存路径。
   - 源文件路径自动清空，焦点回到“浏览”按钮，方便选择下一个文件。

6. **手动截图 PDF**（独立功能）  
   - 点击“📷 手动截图PDF”按钮。
   - 选择 PDF 文件。
   - 选择截图保存目录。
   - 在弹出的图像窗口中，用鼠标拖拽选择区域，按 `Enter` 保存，`n` 下一页，`q` 退出。

## 📁 输出文件说明

- 转换后的文件保存在指定的输出目录中。
- 文件名 = 源文件名（去除扩展名）+ 目标格式扩展名。如果重名，会自动添加数字后缀（如 `文档_1.txt`）。
- **TXT/MD 文件**：只包含纯文本，页面之间用两个换行分隔，无页码标记。
- **其他格式**：通过 Pandoc 生成，保留基本排版。

## ⚠️ 注意事项

- 首次运行会检查并自动安装所需的 Python 包（`pillow`、`PyMuPDF`、`pdfplumber`、`ebooklib`、`beautifulsoup4`、`matplotlib`、`pytesseract`、`pdf2image`），请保持网络畅通。
- 如果 Tesseract 未安装，OCR 功能将不可用，但程序仍可正常运行。
- 如果系统已安装 Pandoc 且添加到 PATH，程序会优先使用系统 Pandoc；否则自动安装 `pypandoc_binary`。
- 手动截图功能需要 `matplotlib` 支持，会自动安装。
- 转换过程中界面会暂时禁用，转换完成后恢复。

## ❓ 常见问题

**Q：为什么转换后输出文件是空的？**  
A：可能是源文件为纯扫描件且 Tesseract 未安装，或提取过程中出错。请检查是否安装了 Tesseract，或尝试将源文件转换为 txt/md 格式后再用其他工具处理。

**Q：如何关闭控制台窗口？**  
A：将脚本扩展名改为 `.pyw`，双击运行即可不显示控制台。

**Q：Pandoc 安装失败怎么办？**  
A：程序会自动安装 `pypandoc_binary`，无需手动安装。如果网络问题导致安装失败，可手动执行 `pip install pypandoc_binary`。

**Q：手动截图时看不到图像窗口？**  
A：检查是否安装了 `matplotlib`，或尝试在命令行运行以查看错误信息。

## 📝 许可证

本工具为自由软件，基于 vibe coding 开发，不可随意修改分发。部分功能依赖开源项目（如 PyMuPDF、pdfplumber、pytesseract 等），请遵守其相应许可证。

---

**仅供学习交流，严禁用于商业用途！** 
