import os
import fitz
import pdfplumber
import ebooklib
from ebooklib import epub
from bs4 import BeautifulSoup
from pathlib import Path
import tempfile
import subprocess
import sys

class DocumentConverter:
    def __init__(self, tesseract_available=False, pandoc_available=False):
        self.tesseract_available = tesseract_available
        self.pandoc_available = pandoc_available

    def extract_text_from_pdf(self, pdf_path, progress_callback=None):
        """从 PDF 提取文本，支持文字型和扫描型（OCR）"""
        text_content = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                for i, page in enumerate(pdf.pages):
                    if progress_callback:
                        progress_callback(f"正在处理 PDF 第 {i+1}/{total_pages} 页...")
                    
                    page_text = page.extract_text()
                    # 如果提取不到文字且 Tesseract 可用，尝试 OCR
                    if (not page_text or len(page_text.strip()) < 10) and self.tesseract_available:
                        page_text = self._ocr_page(pdf_path, i)
                    
                    if page_text:
                        text_content.append(page_text)
            return "\n\n".join(text_content)
        except Exception as e:
            raise Exception(f"PDF 提取失败: {e}")

    def _ocr_page(self, pdf_path, page_index):
        """对单个 PDF 页面进行 OCR"""
        import pytesseract
        from PIL import Image
        import io
        
        doc = fitz.open(pdf_path)
        page = doc[page_index]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2)) # 2x 缩放提高识别率
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        text = pytesseract.image_to_string(img, lang='chi_sim+eng')
        doc.close()
        return text

    def extract_text_from_epub(self, epub_path, progress_callback=None):
        """从 EPUB 提取文本"""
        try:
            book = epub.read_epub(epub_path)
            chapters = []
            items = list(book.get_items_of_type(ebooklib.ITEM_DOCUMENT))
            total = len(items)
            
            for i, item in enumerate(items):
                if progress_callback:
                    progress_callback(f"正在处理 EPUB 章节 {i+1}/{total}...")
                soup = BeautifulSoup(item.get_content(), 'html.parser')
                text = soup.get_text()
                if text.strip():
                    chapters.append(text.strip())
            return "\n\n".join(chapters)
        except Exception as e:
            raise Exception(f"EPUB 提取失败: {e}")

    def convert(self, src_path, dst_format, output_dir, progress_callback=None):
        """执行转换主逻辑"""
        src_path = Path(src_path)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        base_name = src_path.stem
        ext = f".{dst_format}"
        out_file = self._get_unique_path(output_dir, base_name, ext)
        
        if progress_callback:
            progress_callback("正在提取源文件内容...")
            
        # 1. 提取文本
        if src_path.suffix.lower() == '.pdf':
            content = self.extract_text_from_pdf(src_path, progress_callback)
        elif src_path.suffix.lower() == '.epub':
            content = self.extract_text_from_epub(src_path, progress_callback)
        else:
            raise Exception("不支持的文件格式")

        # 2. 根据目标格式保存
        if dst_format in ['txt', 'md']:
            with open(out_file, 'w', encoding='utf-8') as f:
                f.write(content)
            return str(out_file)
        else:
            # 需要 Pandoc 的格式
            if not self.pandoc_available:
                raise Exception("需要 Pandoc 才能转换为该格式")
            
            return self._convert_via_pandoc(content, out_file, dst_format, progress_callback)

    def _convert_via_pandoc(self, content, out_file, dst_format, progress_callback):
        import pypandoc
        temp_txt = tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode='w', encoding='utf-8')
        try:
            temp_txt.write(content)
            temp_txt.close()
            
            if progress_callback:
                progress_callback(f"正在通过 Pandoc 转换为 {dst_format}...")
                
            pypandoc.convert_file(temp_txt.name, dst_format, format='markdown', outputfile=str(out_file))
            return str(out_file)
        finally:
            if os.path.exists(temp_txt.name):
                os.unlink(temp_txt.name)

    def _get_unique_path(self, directory, base, ext):
        counter = 1
        name = f"{base}{ext}"
        full = directory / name
        while full.exists():
            name = f"{base}_{counter}{ext}"
            full = directory / name
            counter += 1
        return full
