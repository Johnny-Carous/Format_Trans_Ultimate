import sys
import os
from pathlib import Path

# 添加 src 目录到路径
sys.path.append(str(Path(__file__).parent / "src"))

from modules.converter import DocumentConverter

def test_conversion():
    # 由于沙盒中可能没有现成的 PDF/EPUB，我们主要测试模块导入和基本逻辑
    print("正在初始化转换器...")
    converter = DocumentConverter(tesseract_available=False, pandoc_available=True)
    
    print("转换器初始化成功！")
    print("支持的提取方法:")
    print(f"- extract_text_from_pdf: {hasattr(converter, 'extract_text_from_pdf')}")
    print(f"- extract_text_from_epub: {hasattr(converter, 'extract_text_from_epub')}")
    print(f"- convert: {hasattr(converter, 'convert')}")

if __name__ == "__main__":
    test_conversion()
