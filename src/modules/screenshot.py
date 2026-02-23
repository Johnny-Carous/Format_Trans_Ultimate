import os
import fitz
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.widgets import RectangleSelector
import numpy as np
from PIL import Image
import io
from pathlib import Path

class PDFScreenshotter:
    def __init__(self, pdf_path, output_dir, dpi=300):
        self.pdf_path = pdf_path
        self.output_dir = Path(output_dir)
        self.dpi = dpi
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.doc = fitz.open(pdf_path)
        self.total_pages = len(self.doc)
        self.page_idx = 0
        self.coords = []
        self.pan_active = False
        self.pan_start = None
        self.pil_img = None

    def start(self):
        """启动截图交互界面"""
        while 0 <= self.page_idx < self.total_pages:
            self._render_page()
            if self.next_action == 'quit':
                break
            elif self.next_action == 'next':
                self.page_idx += 1
            elif self.next_action == 'prev':
                self.page_idx -= 1
            elif self.next_action == 'goto':
                # 这里可以通过 simpledialog 询问页码，但由于在线程中，需要通过回调或队列
                pass
        self.doc.close()

    def _render_page(self):
        page = self.doc[self.page_idx]
        zoom = self.dpi / 72
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        self.pil_img = Image.open(io.BytesIO(img_data))
        img_array = np.array(self.pil_img)

        fig, ax = plt.subplots(figsize=(12, 8))
        ax.imshow(img_array)
        ax.set_title(f"第 {self.page_idx+1}/{self.total_pages} 页 - s:保存, n:下一页, b:上一页, q:退出, r:重置, Ctrl+拖拽:平移, 滚轮:缩放")
        plt.tight_layout()

        self.initial_xlim = ax.get_xlim()
        self.initial_ylim = ax.get_ylim()
        self.next_action = 'quit' # 默认动作

        # 矩形选择器
        def onselect(eclick, erelease):
            x1, y1 = int(eclick.xdata), int(eclick.ydata)
            x2, y2 = int(erelease.xdata), int(erelease.ydata)
            self.coords = [(min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))]
            rect = patches.Rectangle((min(x1, x2), min(y1, y2)), abs(x2-x1), abs(y2-y1), 
                                     linewidth=2, edgecolor='r', facecolor='none')
            ax.add_patch(rect)
            fig.canvas.draw()

        rs = RectangleSelector(ax, onselect, useblit=True, button=[1], minspanx=5, minspany=5, 
                               spancoords='pixels', interactive=True)

        # 滚轮缩放
        def on_scroll(event):
            scale = 1.2 if event.button == 'up' else 0.8
            xlim, ylim = ax.get_xlim(), ax.get_ylim()
            if event.xdata is None or event.ydata is None: return
            new_w = (xlim[1] - xlim[0]) * scale
            new_h = (ylim[1] - ylim[0]) * scale
            ax.set_xlim([event.xdata - new_w/2, event.xdata + new_w/2])
            ax.set_ylim([event.ydata - new_h/2, event.ydata + new_h/2])
            fig.canvas.draw()

        # 键盘事件
        def on_key(event):
            if event.key == 's' and self.coords:
                x1, y1, x2, y2 = self.coords[0]
                cropped = self.pil_img.crop((x1, y1, x2, y2))
                out_path = self._get_unique_path(f"page{self.page_idx+1}_crop.png")
                cropped.save(out_path)
                print(f"已保存: {out_path}")
            elif event.key == 'n':
                self.next_action = 'next'
                plt.close(fig)
            elif event.key == 'b':
                self.next_action = 'prev'
                plt.close(fig)
            elif event.key == 'q':
                self.next_action = 'quit'
                plt.close(fig)
            elif event.key == 'r':
                ax.set_xlim(self.initial_xlim)
                ax.set_ylim(self.initial_ylim)
                fig.canvas.draw()

        fig.canvas.mpl_connect('scroll_event', on_scroll)
        fig.canvas.mpl_connect('key_press_event', on_key)
        plt.show()

    def _get_unique_path(self, filename):
        base, ext = os.path.splitext(filename)
        counter = 1
        full_path = self.output_dir / filename
        while full_path.exists():
            full_path = self.output_dir / f"{base}_{counter}{ext}"
            counter += 1
        return str(full_path)
