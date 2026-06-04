"""批量比对 — 多对文件管理、导航、裁剪复用"""

import os
import io
import tkinter as tk
import threading
import fitz
from docx import Document
from PIL import Image, ImageDraw
from .config import _BaseMixin

class BatchMixin(_BaseMixin):
    """批量比对 — 多对文件管理、导航、裁剪复用"""

    def _on_pair_ocr_done(self, index: int):
        """单对OCR完成：保存数据，若为首对则询问裁剪复用策略"""
        self._save_pair_data(index)
        self._update_batch_nav()
        if index == 0 and len(self._pairs) > 1:
            self._ask_crop_reuse()
        elif index > 0 and getattr(self, '_batch_crop_mode', '') == 'individual':
            self._batch_load_next_with_crop(index + 1)



    def _ask_crop_reuse(self):
        """弹窗询问是否复用首对裁剪区域"""
        win = tk.Toplevel(self.root)
        win.withdraw()
        win.title("批量裁剪策略")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        f = tk.Frame(win, padx=20, pady=15, bg="white")
        f.pack(fill=tk.BOTH, expand=True)
        tk.Label(f, text=f"第1对已处理完成，剩余 {len(self._pairs) - 1} 对。",
                 font=("Microsoft YaHei UI", 10), bg="white").pack(pady=(0, 5))
        tk.Label(f, text="是否将相同的裁剪区域应用于其余文件？",
                 font=("Microsoft YaHei UI", 10, "bold"), bg="white").pack(pady=(0, 12))
        btn_frame = tk.Frame(f, bg="white")
        btn_frame.pack()
        tk.Button(btn_frame, text="是，全部复用",
                  command=lambda: self._on_crop_choice(win, 'reuse'),
                  bg=self.PRIMARY, fg="white", relief=tk.FLAT,
                  padx=16, pady=6, cursor="hand2",
                  font=("Microsoft YaHei UI", 10)).pack(side=tk.LEFT, padx=4)
        tk.Button(btn_frame, text="否，逐对裁剪",
                  command=lambda: self._on_crop_choice(win, 'individual'),
                  bg="#6c757d", fg="white", relief=tk.FLAT,
                  padx=16, pady=6, cursor="hand2",
                  font=("Microsoft YaHei UI", 10)).pack(side=tk.LEFT, padx=4)
        win.bind("<Escape>", lambda _: self._on_crop_choice(win, 'reuse'))
        win.protocol("WM_DELETE_WINDOW", lambda: self._on_crop_choice(win, 'reuse'))
        win.update_idletasks()
        pw, ph = self.root.winfo_width(), self.root.winfo_height()
        px, py = self.root.winfo_x(), self.root.winfo_y()
        ww, wh = win.winfo_width(), win.winfo_height()
        win.geometry(f"+{px + (pw - ww) // 2}+{py + (ph - wh) // 2}")
        win.deiconify()
        win.wait_window()



    def _on_crop_choice(self, win, mode: str):
        self._batch_crop_mode = mode
        win.destroy()
        if mode == 'reuse':
            self._batch_load_next(1)
        else:
            self._batch_load_next_with_crop(1)



    def _save_pair_data(self, index: int):
        self._pair_data[index] = {
            'docx_text': self.docx_text,
            'pdf_text': self.pdf_text,
            'doc_obj': self.doc_obj,
            'cleaned': self.cleaned_pdf_images,
            'docx_path': self.docx_path,
            'pdf_path': self.pdf_path,
            'pdf_name': self._pdf_name,
            'docx_paragraphs': self._docx_paragraphs,
            'diff_blocks': self.diff_blocks,
            'docx_flat_positions': self._docx_flat_positions,
            'docx_flat_to_raw': getattr(self, '_docx_flat_to_raw', []),
            'saved_docx_source': getattr(self, '_saved_docx_source', ''),
            'docx_text_raw': getattr(self, '_docx_text_raw', ''),
        }



    def _update_batch_nav(self):
        total = len(self._pairs)
        loaded = len(self._pair_data)
        self._batch_lbl.config(text=f"第 {self._pair_index + 1}/{total} 对（已加载 {loaded} 对）")
        self._update_button_states()



    def _batch_load_next(self, index: int):
        """后台加载下一对（跳过裁剪对话框，复用首对 discard_boxes）"""
        if index >= len(self._pairs):
            return
        docx_path, pdf_path = self._pairs[index]
        import threading

        def worker():
            try:
                # 加载 DOCX
                doc = Document(docx_path)
                paras = []
                for p in doc.paragraphs:
                    if not self._is_red_header(p) and p.text.strip():
                        paras.append(p.text)
                docx_text = self._normalize_text("\n".join(paras))

                # 加载 PDF 页面
                pdf_doc = fitz.open(pdf_path)
                mat = fitz.Matrix(1.5, 1.5)
                images = []
                for i in range(len(pdf_doc)):
                    page = pdf_doc.load_page(i)
                    pix = page.get_pixmap(matrix=mat)
                    img_data = pix.tobytes("png")
                    img = Image.open(io.BytesIO(img_data)).convert("RGB")
                    cleaned = self._remove_red_pixels(img)
                    # 应用首对的裁剪区域
                    boxes = self.discard_boxes[i % 2]
                    if boxes:
                        draw = ImageDraw.Draw(cleaned)
                        for box in boxes:
                            draw.rectangle(box, fill="white")
                    images.append(cleaned)

                # OCR
                pdf_lines = []
                for img in images:
                    lines = self._ocr_single(img)
                    pdf_lines.extend(lines)
                pdf_text = self._normalize_text("\n".join(pdf_lines))

                self._pair_data[index] = {
                    'docx_text': docx_text,
                    'pdf_text': pdf_text,
                    'doc_obj': doc,
                    'cleaned': images,
                    'docx_path': docx_path,
                    'pdf_path': pdf_path,
                    'pdf_name': os.path.basename(pdf_path),
                    'docx_paragraphs': paras,
                    'docx_text_raw': '\n'.join(paras),
                }
            except Exception as e:
                self._pair_data[index] = {'error': str(e)}

        def poll():
            if thread.is_alive():
                self.root.after(200, poll)
            else:
                if 'error' in self._pair_data.get(index, {}):
                    err = self._pair_data[index]['error']
                    self._alert("批量加载错误",
                                f"第{index + 1}对加载失败:\n{err}", "error")
                self._update_batch_nav()
                # 继续加载下一对
                self._batch_load_next(index + 1)

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()
        self.root.after(100, poll)



    def _batch_load_next_with_crop(self, index: int):
        """逐对加载（含裁剪对话框）"""
        if index >= len(self._pairs):
            self._update_batch_nav()
            return
        docx_path, pdf_path = self._pairs[index]
        # 保存当前对的数据（如果还在激活对的话）
        if self._pair_index >= 0 and self._pair_index not in self._pair_data:
            self._save_pair_data(self._pair_index)
        self._load_docx_path(docx_path)
        self._load_pdf_path(pdf_path, pair_index=index)



    def _switch_to_pair(self, index: int):
        """切换到指定索引的比对对"""
        if index not in self._pair_data:
            self._set_status(f"第{index + 1}对尚未加载完成，请稍候", "warn")
            return
        data = self._pair_data[index]
        self._pair_index = index
        self.docx_text = data['docx_text']
        self.pdf_text = data['pdf_text']
        self.doc_obj = data['doc_obj']
        self.cleaned_pdf_images = data['cleaned']
        self.docx_path = data['docx_path']
        self.pdf_path = data['pdf_path']
        self._pdf_name = data['pdf_name']
        self._docx_paragraphs = data.get('docx_paragraphs', [])
        self._docx_flat_positions = data.get('docx_flat_positions', [])
        self._docx_flat_to_raw = data.get('docx_flat_to_raw', [])
        self._saved_docx_source = data.get('saved_docx_source', '')
        self._docx_text_raw = data.get('docx_text_raw', '')
        saved_blocks = data.get('diff_blocks', [])
        self.diff_blocks = saved_blocks

        self.lbl_docx.config(text=os.path.basename(self.docx_path or ""), fg=self.PRIMARY)
        self.lbl_pdf.config(text=f"{self._pdf_name}（OCR完成）", fg="green")
        self._update_batch_nav()

        # 已有已保存的差异块则直接渲染，否则重新分析
        if saved_blocks:
            self._hide_hover_popup()
            self._render_diff()
            self.txt_diff.bind("<Motion>", self._on_diff_motion)
            self._update_button_states()
        elif self.txt_diff.get("1.0", tk.END).strip():
            self.analyze_diff(show_warning=False)



    def _nav_prev(self):
        if self._pair_index > 0:
            self._save_pair_data(self._pair_index)
            self._switch_to_pair(self._pair_index - 1)



    def _nav_next(self):
        if self._pair_index < len(self._pairs) - 1:
            self._save_pair_data(self._pair_index)
            self._switch_to_pair(self._pair_index + 1)

