import os
import sys
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
from modules.converter import DocumentConverter
from modules.screenshot import PDFScreenshotter

# 设置外观
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("文档格式转换工具 Ultimate")
        self.geometry("700x500")
        
        # 初始化变量
        self.src_path = ctk.StringVar()
        self.dst_format = ctk.StringVar(value="txt")
        self.output_dir = ctk.StringVar(value=str(Path.cwd() / "output"))
        self.status_text = ctk.StringVar(value="就绪")
        
        self.tesseract_available = False
        self.pandoc_available = False
        
        self._setup_ui()
        self._check_dependencies()
        
        self.converter = DocumentConverter(self.tesseract_available, self.pandoc_available)

    def _setup_ui(self):
        # 配置网格
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 侧边栏
        self.sidebar_frame = ctk.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="文件转换器", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        self.sidebar_button_1 = ctk.CTkButton(self.sidebar_frame, text="手动截图 PDF", command=self.manual_crop)
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        
        self.appearance_mode_label = ctk.CTkLabel(self.sidebar_frame, text="外观模式:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
        self.appearance_mode_optionemenu.set("System")

        # 主内容区
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.grid(row=0, column=1, padx=(20, 20), pady=(20, 20), sticky="nsew")
        self.main_frame.grid_columnconfigure(1, weight=1)

        # 源文件选择
        self.label_src = ctk.CTkLabel(self.main_frame, text="源文件:")
        self.label_src.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        self.entry_src = ctk.CTkEntry(self.main_frame, textvariable=self.src_path, placeholder_text="请选择 PDF 或 EPUB 文件")
        self.entry_src.grid(row=0, column=1, padx=(20, 10), pady=(20, 10), sticky="ew")
        self.btn_browse_src = ctk.CTkButton(self.main_frame, text="浏览", width=80, command=self.browse_src)
        self.btn_browse_src.grid(row=0, column=2, padx=(10, 20), pady=(20, 10))

        # 目标格式
        self.label_fmt = ctk.CTkLabel(self.main_frame, text="目标格式:")
        self.label_fmt.grid(row=1, column=0, padx=20, pady=10, sticky="w")
        self.fmt_menu = ctk.CTkOptionMenu(self.main_frame, values=["txt", "md", "docx", "pdf", "epub"], variable=self.dst_format)
        self.fmt_menu.grid(row=1, column=1, padx=20, pady=10, sticky="w")

        # 输出目录
        self.label_out = ctk.CTkLabel(self.main_frame, text="输出目录:")
        self.label_out.grid(row=2, column=0, padx=20, pady=10, sticky="w")
        self.entry_out = ctk.CTkEntry(self.main_frame, textvariable=self.output_dir)
        self.entry_out.grid(row=2, column=1, padx=(20, 10), pady=10, sticky="ew")
        self.btn_browse_out = ctk.CTkButton(self.main_frame, text="浏览", width=80, command=self.browse_output)
        self.btn_browse_out.grid(row=2, column=2, padx=(10, 20), pady=10)

        # 转换按钮
        self.btn_convert = ctk.CTkButton(self.main_frame, text="开始转换", font=ctk.CTkFont(size=16, weight="bold"), height=40, command=self.start_conversion)
        self.btn_convert.grid(row=3, column=0, columnspan=3, padx=20, pady=30, sticky="ew")

        # 状态和进度
        self.status_label = ctk.CTkLabel(self.main_frame, textvariable=self.status_text, text_color="gray")
        self.status_label.grid(row=4, column=0, columnspan=3, padx=20, pady=(10, 0))
        
        self.progressbar = ctk.CTkProgressBar(self.main_frame)
        self.progressbar.grid(row=5, column=0, columnspan=3, padx=20, pady=(10, 20), sticky="ew")
        self.progressbar.set(0)

    def _check_dependencies(self):
        """检查外部工具依赖"""
        import shutil
        self.tesseract_available = bool(shutil.which('tesseract'))
        try:
            import pypandoc
            pypandoc.get_pandoc_version()
            self.pandoc_available = True
        except:
            self.pandoc_available = False
            
        if not self.tesseract_available:
            print("警告: 未找到 Tesseract，OCR 功能将不可用。")
        if not self.pandoc_available:
            print("警告: 未找到 Pandoc，部分格式转换将不可用。")

    def browse_src(self):
        path = filedialog.askopenfilename(filetypes=[("文档文件", "*.pdf *.epub"), ("所有文件", "*.*")])
        if path:
            self.src_path.set(path)

    def browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)

    def start_conversion(self):
        src = self.src_path.get().strip()
        if not src:
            messagebox.showerror("错误", "请选择源文件")
            return
        
        self.btn_convert.configure(state="disabled")
        self.progressbar.configure(mode="indeterminate")
        self.progressbar.start()
        
        thread = threading.Thread(target=self._convert_worker, args=(src, self.dst_format.get(), self.output_dir.get()))
        thread.daemon = True
        thread.start()

    def _convert_worker(self, src, fmt, out_dir):
        try:
            def update_status(msg):
                self.after(0, lambda: self.status_text.set(msg))
            
            result_path = self.converter.convert(src, fmt, out_dir, progress_callback=update_status)
            self.after(0, lambda: self._on_conversion_success(result_path))
        except Exception as e:
            self.after(0, lambda: self._on_conversion_error(str(e)))

    def _on_conversion_success(self, path):
        self.progressbar.stop()
        self.progressbar.configure(mode="determinate")
        self.progressbar.set(1)
        self.btn_convert.configure(state="normal")
        self.status_text.set("转换成功！")
        messagebox.showinfo("完成", f"文件已保存至：\n{path}")
        self.src_path.set("")

    def _on_conversion_error(self, error_msg):
        self.progressbar.stop()
        self.progressbar.set(0)
        self.btn_convert.configure(state="normal")
        self.status_text.set("转换失败")
        messagebox.showerror("错误", f"转换失败：{error_msg}")

    def manual_crop(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path: return
        
        out_dir = filedialog.askdirectory(title="选择截图保存目录")
        if not out_dir: out_dir = str(Path(path).parent / "crops")
        
        def run_screenshot():
            try:
                sc = PDFScreenshotter(path, out_dir)
                sc.start()
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("错误", f"截图工具启动失败: {e}"))

        thread = threading.Thread(target=run_screenshot)
        thread.daemon = True
        thread.start()
        self.status_text.set("截图工具已启动...")

if __name__ == "__main__":
    app = App()
    app.mainloop()
