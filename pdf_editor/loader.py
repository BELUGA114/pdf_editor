"""文件加载 — DOCX/PDF 加载、文本提取、文件扫描"""

import os
import io
import tkinter as tk
from tkinter import filedialog
import numpy as np
import fitz
from docx import Document
from PIL import Image, ImageDraw
from .config import _BaseMixin

class LoaderMixin(_BaseMixin):
    """文件加载 — DOCX/PDF 加载、文本提取、文件扫描"""

    def _is_red_header(self, paragraph):
        """检查段落中的文字是否包含红色字体"""
        for run in paragraph.runs:
            if run.font.color and run.font.color.rgb:
                r, g, b = run.font.color.rgb
                if r > 200 and g < 60 and b < 60:
                    return True
        return False



    def _load_docx_path(self, path: str):
        """从路径加载 DOCX（支持拖放）"""
        # 非批量模式：清除批量状态
        if not self._pairs:
            self._batch_nav_frame.pack_forget()
            self._pair_data = {}
            self._pair_index = -1
        self.docx_path = path
        self.lbl_docx.config(text=os.path.basename(path), fg=self.PRIMARY)
        self.doc_obj = Document(path)
        body_paragraphs = []
        self._docx_paragraphs = []
        for p in self.doc_obj.paragraphs:
            if self._is_red_header(p):
                continue
            if p.text.strip():
                body_paragraphs.append(p.text)
                self._docx_paragraphs.append(p.text)
        raw_text = "\n".join(body_paragraphs)
        self._docx_text_raw = raw_text
        self.docx_text = self._normalize_text(raw_text)
        self.diff_blocks = []



    def load_docx(self):
        """通过文件对话框加载 DOCX"""
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path:
            self._load_docx_path(path)



    def _scan_folder(self, path: str, exts: set):
        """扫描文件夹中指定扩展名的文件，返回 {stem: fullpath}"""
        result = {}
        for f in os.listdir(path):
            full = os.path.join(path, f)
            if not os.path.isfile(full):
                continue
            stem, ext = os.path.splitext(f)
            if ext.lower() in exts:
                result[stem] = full
        return result



    def _try_folder_match(self):
        """当两个文件夹都选定后，自动按文件名匹配并加载（支持批量）"""
        if not self._docx_folder or not self._pdf_folder:
            return

        docx_files = self._scan_folder(self._docx_folder, {".docx", ".doc"})
        pdf_files = self._scan_folder(self._pdf_folder, {".pdf"})

        pairs = []
        for stem in docx_files:
            if stem in pdf_files:
                pairs.append((docx_files[stem], pdf_files[stem]))

        if not pairs:
            self._alert("提示",
                f"两个文件夹中未找到文件名匹配的 DOCX/PDF 对。\n"
                f"DOCX 文件夹: {len(docx_files)} 个\n"
                f"PDF 文件夹: {len(pdf_files)} 个",
                "warn")
            return

        self._pairs = pairs
        self._pair_index = 0
        self._pair_data = {}

        # 显示批量导航
        if len(pairs) > 1:
            self._batch_nav_frame.pack(fill=tk.X, pady=(0, 5))

        docx_path, pdf_path = pairs[0]
        self._load_docx_path(docx_path)
        self._load_pdf_path(pdf_path, pair_index=0)
        if len(pairs) > 1:
            self._alert("提示",
                f"找到 {len(pairs)} 对匹配文件，请先裁剪第1对PDF。\n"
                f"完成后可选择是否复用裁剪区域。",
                "info")



    def _select_docx_folder(self):
        path = filedialog.askdirectory(title="选择存放 DOCX 文件的文件夹")
        if path:
            self._docx_folder = path
            self.lbl_docx_folder.config(text=os.path.basename(path), fg=self.PRIMARY)
            self._try_folder_match()



    def _select_pdf_folder(self):
        path = filedialog.askdirectory(title="选择存放 PDF 文件的文件夹")
        if path:
            self._pdf_folder = path
            self.lbl_pdf_folder.config(text=os.path.basename(path), fg=self.PRIMARY)
            self._try_folder_match()

    @staticmethod


    def _normalize_text(text):
        """归一化文本：全角转半角、合并连续空格、去除行首行尾空白"""
        result = []
        for ch in text:
            code = ord(ch)
            if 0xFF01 <= code <= 0xFF5E:  # 全角ASCII → 半角
                result.append(chr(code - 0xFEE0))
            elif code == 0x3000:  # 全角空格 → 半角空格
                result.append(' ')
            else:
                result.append(ch)
        normalized = ''.join(result)
        # 合并连续空格，逐行strip
        lines = []
        for line in normalized.splitlines():
            cleaned = ' '.join(line.split())
            if cleaned:
                lines.append(cleaned)
        return '\n'.join(lines)


    @staticmethod
    def _remove_red_pixels(img):
        """将偏红色像素替换为白色（numpy 向量化，比逐像素循环快数十倍）"""
        arr = np.array(img)
        r, g, b = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2]
        mask = r > np.maximum(g, b) + 30
        arr[mask] = [255, 255, 255]
        return Image.fromarray(arr)

    def _load_pdf_path(self, path: str, pair_index: int | None = None):
        """从路径加载 PDF（支持拖放），包括去红头 + 裁剪 → OCR"""
        self.pdf_path = path
        self._pdf_name = os.path.basename(path)
        self._set_status(f"正在加载 {self._pdf_name} ...", self.WARN)

        self.cleaned_pdf_images = []
        self.crop_box = None
        self._current_pair_index = pair_index
        try:
            doc = fitz.open(path)
            mat = fitz.Matrix(1.5, 1.5)
            for i in range(len(doc)):
                page = doc.load_page(i)
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data)).convert("RGB")
                cleaned = self._remove_red_pixels(img)
                self.cleaned_pdf_images.append(cleaned)
                self._set_status(f"加载中... 第 {i + 1}/{len(doc)} 页")

            self.lbl_pdf.config(text=f"{self._pdf_name}（已加载，请裁剪）", fg=self.WARN)
            self._show_crop_dialog()
        except Exception as e:
            self.lbl_pdf.config(text=f"{self._pdf_name}（加载失败）", fg="red")
            self._set_status("")
            self._alert("错误", f"加载PDF异常:\n{str(e)}", "error")



    def load_pdf(self):
        """通过文件对话框加载 PDF"""
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self._load_pdf_path(path)

