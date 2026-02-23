#!/usr/bin/env python3
"""
    文档格式转换工具 Ultimate v3
    优化内容：
    1. UI 界面：使用 customtkinter 现代 UI 库，支持深色/浅色模式，布局更清晰。
    2. 功能重构：模块化设计，异步多线程处理，防止界面卡死。
    3. 增强功能：优化了 PDF/EPUB 提取逻辑，改进了截图工具的交互体验。
"""

import os
import sys
import threading
import subprocess
import tempfile
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import warnings

# 忽略警告
warnings.filterwarnings("ignore")

# ------------------ 核心依赖检查和自动安装 ------------------
def ensure_pip_package(module_name, package_name=None):
    if package_name is None: package_name = module_name
    try:
        __import__(module_name)
        return True
    except ImportError:
        print(f"正在安装 {package_name}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            __import__(module_name)
            return True
        except:
            return False

REQUIRED_PACKAGES = {
    "customtkinter": "customtkinter",
    "PIL": "pillow",
    "fitz": "PyMuPDF",
    "pdfplumber": "pdfplumber",
    "ebooklib": "ebooklib",
    "bs4": "beautifulsoup4",
    "matplotlib": "matplotlib",
    "pytesseract": "pytesseract",
    "pdf2image": "pdf2image",
    "numpy": "numpy"
}

for mod, pkg in REQUIRED_PACKAGES.items():
    ensure_pip_package(mod, pkg)

import customtkinter as ctk
import fitz
import pdfplumber
from ebooklib import epub
from bs4 import BeautifulSoup
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.widgets import RectangleSelector
import numpy as np
from PIL import Image
import io

# ------------------ 核心逻辑类 ------------------

class DocumentConverter:
    def __init__(self, tesseract_available=False, pandoc_available=False):
        self.tesseract_available = tesseract_available
        self.pandoc_available = pandoc_available

    def extract_text_from_pdf(self, pdf_path, progress_callback=None):
        text_content = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                for i, page in enumerate(pdf.pages):
                    if progress_callback: progress_callback(f"正在处理 PDF 第 {i+1}/{total_pages} 页...")
                    page_text = page.extract_text()
                    if (not page_text or len(page_text.strip()) < 10) and self.tesseract_available:
                        page_text = self._ocr_page(pdf_path, i)
                    if page_text: text_content.append(page_text)
            return "\n\n".join(text_content)
        except Exception as e:
            raise Exception(f"PDF 提取失败: {e}")

    def _ocr_page(self, pdf_path, page_index):
        import pytesseract
        doc = fitz.open(pdf_path)
        page = doc[page_index]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        text = pytesseract.image_to_string(img, lang='chi_sim+eng')
        doc.close()
        return text

    def extract_text_from_epub(self, epub_path, progress_callback=None):
        try:
            book = epub.read_epub(epub_path)
            chapters = []
            items = list(book.get_items_of_type(1)) # ebooklib.ITEM_DOCUMENT
            total = len(items)
            for i, item in enumerate(items):
                if progress_callback: progress_callback(f"正在处理 EPUB 章节 {i+1}/{total}...")
                soup = BeautifulSoup(item.get_content(), 'html.parser')
                text = soup.get_text()
                if text.strip(): chapters.append(text.strip())
            return "\n\n".join(chapters)
        except Exception as e:
            raise Exception(f"EPUB 提取失败: {e}")

    def convert(self, src_path, dst_format, output_dir, progress_callback=None):
        src_path = Path(src_path)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        out_file = self._get_unique_path(output_dir, src_path.stem, f".{dst_format}")
        
        if src_path.suffix.lower() == '.pdf':
            content = self.extract_text_from_pdf(src_path, progress_callback)
        elif src_path.suffix.lower() == '.epub':
            content = self.extract_text_from_epub(src_path, progress_callback)
        else:
            raise Exception("不支持的文件格式")

        if dst_format in ['txt', 'md']:
            with open(out_file, 'w', encoding='utf-8') as f:
                f.write(content)
            return str(out_file)
        else:
            if not self.pandoc_available: raise Exception("需要 Pandoc 才能转换为该格式")
            return self._convert_via_pandoc(content, out_file, dst_format, progress_callback)

    def _convert_via_pandoc(self, content, out_file, dst_format, progress_callback):
        import pypandoc
        temp_txt = tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode='w', encoding='utf-8')
        try:
            temp_txt.write(content)
            temp_txt.close()
            if progress_callback: progress_callback(f"正在通过 Pandoc 转换为 {dst_format}...")
            pypandoc.convert_file(temp_txt.name, dst_format, format='markdown', outputfile=str(out_file))
            return str(out_file)
        finally:
            if os.path.exists(temp_txt.name): os.unlink(temp_txt.name)

    def _get_unique_path(self, directory, base, ext):
        counter = 1
        name = f"{base}{ext}"
        full = directory / name
        while full.exists():
            name = f"{base}_{counter}{ext}"
            full = directory / name
            counter += 1
        return full

# ------------------ GUI 界面类 ------------------

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("文档格式转换工具 Ultimate v3")
        self.geometry("750x550")
        
        self.src_path = tk.StringVar()
        self.dst_format = tk.StringVar(value="txt")
        self.output_dir = tk.StringVar(value=str(Path.cwd() / "output"))
        self.status_text = tk.StringVar(value="就绪")
        
        self.tesseract_available = False
        self.pandoc_available = False
        
        self._setup_ui()
        self._check_dependencies()
        self.converter = DocumentConverter(self.tesseract_available, self.pandoc_available)

    def _setup_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 侧边栏
        self.sidebar_frame = ctk.CTkFrame(self, width=160, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="文件转换器", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        self.btn_manual = ctk.CTkButton(self.sidebar_frame, text="📷 手动截图 PDF", command=self.manual_crop)
        self.btn_manual.grid(row=1, column=0, padx=20, pady=10)
        
        self.appearance_mode_label = ctk.CTkLabel(self.sidebar_frame, text="外观模式:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_menu = ctk.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"], command=lambda m: ctk.set_appearance_mode(m))
        self.appearance_mode_menu.grid(row=6, column=0, padx=20, pady=(10, 20))
        self.appearance_mode_menu.set("System")

        # 主内容区
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(1, weight=1)

        # 源文件
        ctk.CTkLabel(self.main_frame, text="源文件:").grid(row=0, column=0, padx=20, pady=(30, 10), sticky="w")
        self.entry_src = ctk.CTkEntry(self.main_frame, textvariable=self.src_path, placeholder_text="请选择 PDF 或 EPUB 文件")
        self.entry_src.grid(row=0, column=1, padx=(10, 10), pady=(30, 10), sticky="ew")
        ctk.CTkButton(self.main_frame, text="浏览", width=80, command=self.browse_src).grid(row=0, column=2, padx=(10, 20), pady=(30, 10))

        # 格式
        ctk.CTkLabel(self.main_frame, text="目标格式:").grid(row=1, column=0, padx=20, pady=10, sticky="w")
        self.fmt_menu = ctk.CTkOptionMenu(self.main_frame, values=["txt", "md", "docx", "pdf", "epub"], variable=self.dst_format)
        self.fmt_menu.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        # 输出
        ctk.CTkLabel(self.main_frame, text="输出目录:").grid(row=2, column=0, padx=20, pady=10, sticky="w")
        self.entry_out = ctk.CTkEntry(self.main_frame, textvariable=self.output_dir)
        self.entry_out.grid(row=2, column=1, padx=(10, 10), pady=10, sticky="ew")
        ctk.CTkButton(self.main_frame, text="浏览", width=80, command=self.browse_output).grid(row=2, column=2, padx=(10, 20), pady=10)

        # 转换
        self.btn_convert = ctk.CTkButton(self.main_frame, text="开始转换", font=ctk.CTkFont(size=16, weight="bold"), height=45, command=self.start_conversion)
        self.btn_convert.grid(row=3, column=0, columnspan=3, padx=20, pady=40, sticky="ew")

        # 状态
        self.status_label = ctk.CTkLabel(self.main_frame, textvariable=self.status_text, text_color="gray")
        self.status_label.grid(row=4, column=0, columnspan=3, padx=20, pady=(10, 0))
        self.progressbar = ctk.CTkProgressBar(self.main_frame)
        self.progressbar.grid(row=5, column=0, columnspan=3, padx=20, pady=(10, 20), sticky="ew")
        self.progressbar.set(0)

    def _check_dependencies(self):
        self.tesseract_available = bool(shutil.which('tesseract'))
        try:
            import pypandoc
            pypandoc.get_pandoc_version()
            self.pandoc_available = True
        except:
            self.pandoc_available = False

    def browse_src(self):
        path = filedialog.askopenfilename(filetypes=[("文档文件", "*.pdf *.epub"), ("所有文件", "*.*")])
        if path: self.src_path.set(path)

    def browse_output(self):
        path = filedialog.askdirectory()
        if path: self.output_dir.set(path)

    def start_conversion(self):
        src = self.src_path.get().strip()
        if not src:
            messagebox.showerror("错误", "请选择源文件")
            return
        self.btn_convert.configure(state="disabled")
        self.progressbar.configure(mode="indeterminate")
        self.progressbar.start()
        threading.Thread(target=self._convert_worker, args=(src, self.dst_format.get(), self.output_dir.get()), daemon=True).start()

    def _convert_worker(self, src, fmt, out_dir):
        try:
            def update_status(msg): self.after(0, lambda: self.status_text.set(msg))
            result_path = self.converter.convert(src, fmt, out_dir, progress_callback=update_status)
            self.after(0, lambda: self._on_done(True, result_path))
        except Exception as e:
            self.after(0, lambda: self._on_done(False, str(e)))

    def _on_done(self, success, result):
        self.progressbar.stop()
        self.progressbar.configure(mode="determinate")
        self.progressbar.set(1 if success else 0)
        self.btn_convert.configure(state="normal")
        if success:
            self.status_text.set("转换成功！")
            messagebox.showinfo("完成", f"文件已保存至：\n{result}")
            self.src_path.set("")
        else:
            self.status_text.set("转换失败")
            messagebox.showerror("错误", f"转换失败：{result}")

    def manual_crop(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path: return
        out_dir = filedialog.askdirectory(title="选择截图保存目录")
        if not out_dir: out_dir = str(Path(path).parent / "crops")
        
        def run_screenshot():
            try:
                doc = fitz.open(path)
                page_idx = 0
                while 0 <= page_idx < len(doc):
                    page = doc[page_idx]
                    pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
                    pil_img = Image.open(io.BytesIO(pix.tobytes("png")))
                    
                    fig, ax = plt.subplots(figsize=(12, 8))
                    ax.imshow(np.array(pil_img))
                    ax.set_title(f"第 {page_idx+1}/{len(doc)} 页 - s:保存, n:下一页, b:上一页, q:退出")
                    
                    action = {'type': 'quit'}
                    coords = []
                    def onselect(e1, e2):
                        coords.clear()
                        coords.append((min(e1.xdata, e2.xdata), min(e1.ydata, e2.ydata), max(e1.xdata, e2.xdata), max(e1.ydata, e2.ydata)))
                    
                    rs = RectangleSelector(ax, onselect, useblit=True, button=[1], interactive=True)
                    
                    def on_key(event):
                        if event.key == 's' and coords:
                            x1, y1, x2, y2 = coords[0]
                            cropped = pil_img.crop((x1, y1, x2, y2))
                            out_path = Path(out_dir) / f"crop_{page_idx+1}_{len(os.listdir(out_dir))}.png"
                            cropped.save(out_path)
                            print(f"已保存: {out_path}")
                        elif event.key == 'n': action['type'] = 'next'; plt.close(fig)
                        elif event.key == 'b': action['type'] = 'prev'; plt.close(fig)
                        elif event.key == 'q': action['type'] = 'quit'; plt.close(fig)
                    
                    fig.canvas.mpl_connect('key_press_event', on_key)
                    plt.show()
                    if action['type'] == 'quit': break
                    elif action['type'] == 'next': page_idx += 1
                    elif action['type'] == 'prev': page_idx -= 1
                doc.close()
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("错误", f"截图工具出错: {e}"))

        threading.Thread(target=run_screenshot, daemon=True).start()
        self.status_text.set("截图工具已启动...")

if __name__ == "__main__":
    App().mainloop()
