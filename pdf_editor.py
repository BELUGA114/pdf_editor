import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import pytesseract
from docx import Document
import difflib
import os
import io
from PIL import Image, ImageTk, ImageDraw

# ================= 配置区域 =================
TESSERACT_PATH = r"E:\Code\Tesseract\tesseract.exe"
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
# ============================================


class DocxPdfReviewer:
    def __init__(self, root):
        self.root = root
        self.root.title("Docx-PDF 差异比对与格式化校对工具")
        self.root.geometry("1200x800")

        self.docx_path = None
        self.pdf_path = None
        self.doc_obj = None

        self.docx_text = ""
        self.pdf_text = ""
        self.diff_results = []
        self.cleaned_pdf_images = []

        self._build_ui()

    def _build_ui(self):
        # 顶部控制栏
        top_frame = tk.Frame(self.root, pady=10, padx=10, bg="#f5f5f5")
        top_frame.pack(side=tk.TOP, fill=tk.X)

        tk.Button(top_frame, text="1. 加载原始Docx", command=self.load_docx, width=15).pack(side=tk.LEFT, padx=5)
        self.lbl_docx = tk.Label(top_frame, text="未加载", fg="gray", bg="#f5f5f5")
        self.lbl_docx.pack(side=tk.LEFT, padx=5)

        tk.Button(top_frame, text="2. 加载对比PDF", command=self.load_pdf, width=15).pack(side=tk.LEFT, padx=15)
        self.lbl_pdf = tk.Label(top_frame, text="未加载", fg="gray", bg="#f5f5f5")
        self.lbl_pdf.pack(side=tk.LEFT, padx=5)

        tk.Button(top_frame, text="3. 分析差异 (Git模式)", command=self.analyze_diff, width=18, bg="#d1ecf1").pack(side=tk.LEFT, padx=20)
        tk.Button(top_frame, text="预览去红头PDF", command=self.preview_cleaned_pdf, width=14, bg="#f0e6d3").pack(side=tk.LEFT, padx=5)

        # 导出按钮组
        export_frame = tk.Frame(top_frame, bg="#f5f5f5")
        export_frame.pack(side=tk.RIGHT, padx=5)
        tk.Button(export_frame, text="导出：同步更新版Docx", command=self.save_synced_docx, bg="#d4edda", width=20).pack(side=tk.TOP, pady=2)
        tk.Button(export_frame, text="导出：自动去红头版Docx", command=self.save_dered_docx, bg="#fff3cd", width=20).pack(side=tk.TOP, pady=2)

        # 中部差异展示区（类似Git Diff效果）
        mid_frame = tk.Frame(self.root, padx=10, pady=10)
        mid_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(mid_frame, text="红色=PDF缺少的文字  绿色=PDF多出的文字  （忽略标点符号和格式差异）", font=("Consolas", 10)).pack(anchor=tk.W)

        self.txt_diff = tk.Text(mid_frame, font=("Consolas", 11), wrap=tk.WORD)
        self.txt_diff.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll = tk.Scrollbar(mid_frame, command=self.txt_diff.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_diff.config(yscrollcommand=scroll.set)

        # 差异文本颜色标签配置
        self.txt_diff.tag_config("deletion", background="#ffeef0", foreground="#b30000")  # Git 红
        self.txt_diff.tag_config("addition", background="#e6ffed", foreground="#008000")  # Git 绿
        self.txt_diff.tag_config("header", foreground="#6a737d", font=("Consolas", 10, "italic"))

        # 右键复制菜单
        self._setup_copy_menu(self.txt_diff)

    def _setup_copy_menu(self, widget):
        """为Text组件添加右键复制菜单"""
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="复制", command=lambda: self._copy_selection(widget))
        widget.bind("<Button-3>", lambda _: menu.post(_.x_root, _.y_root))
        widget.bind("<Control-c>", lambda _: self._copy_selection(widget))

    @staticmethod
    def _copy_selection(widget):
        try:
            selected = widget.selection_get()
            widget.clipboard_clear()
            widget.clipboard_append(selected)
        except tk.TclError:
            pass

    def load_docx(self):
        self.docx_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if self.docx_path:
            self.lbl_docx.config(text=os.path.basename(self.docx_path), fg="black")
            self.doc_obj = Document(self.docx_path)
            raw_text = "\n".join([p.text for p in self.doc_obj.paragraphs if p.text.strip()])
            self.docx_text = self._normalize_text(raw_text)

    @staticmethod
    def _remove_red_pixels(img):
        """将偏红色像素替换为白色（相对差值法，适应扫描件颜色不纯）"""
        pixels = img.load()
        width, height = img.size
        for y in range(height):
            for x in range(width):
                r, g, b = pixels[x, y]
                if r > max(g, b) + 30:
                    pixels[x, y] = (255, 255, 255)
        return img

    @staticmethod
    def _strip_whitespace(text):
        """去除所有空白字符（空格、换行、制表符等），仅保留可见文字"""
        return ''.join(ch for ch in text if not ch.isspace())

    @staticmethod
    def _keep_only_words(text):
        """仅保留中英文文字和数字，去除标点符号"""
        return ''.join(
            ch for ch in text
            if '一' <= ch <= '鿿'    # 中文
            or '㐀' <= ch <= '䶿'    # 中文扩展
            or 'a' <= ch <= 'z' or 'A' <= ch <= 'Z'
            or '0' <= ch <= '9'
        )

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

    def load_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not self.pdf_path:
            return

        # 阶段1：加载PDF并去红头（暂不OCR）
        self._pdf_name = os.path.basename(self.pdf_path)
        self.lbl_pdf.config(text=f"正在加载 {self._pdf_name} ...", fg="orange")
        self.root.update()

        self.cleaned_pdf_images = []
        self.crop_box = None
        try:
            doc = fitz.open(self.pdf_path)
            mat = fitz.Matrix(2.0, 2.0)
            for i in range(len(doc)):
                page = doc.load_page(i)
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data)).convert("RGB")
                cleaned = self._remove_red_pixels(img)
                self.cleaned_pdf_images.append(cleaned)
                self.lbl_pdf.config(text=f"加载中... 第 {i + 1}/{len(doc)} 页")
                self.root.update()

            self.lbl_pdf.config(text=f"{self._pdf_name}（已加载，请裁剪）", fg="orange")
            self.root.update()
            # 阶段2：弹出裁剪对话框（首页框选，应用到所有页）
            self._show_crop_dialog()
        except Exception as e:
            self.lbl_pdf.config(text=f"{self._pdf_name}（加载失败）", fg="red")
            messagebox.showerror("错误", f"加载PDF异常:\n{str(e)}")

    def _show_crop_dialog(self):
        """两步裁剪：第1页（单数页基准）→ 第2页（双数页基准）→ OCR"""
        if not self.cleaned_pdf_images:
            return

        self.discard_boxes: list = [None, None]
        win = tk.Toplevel(self.root)
        win.title("裁剪PDF（框选要丢弃的区域）")
        win.geometry("900x700")

        info_lbl = tk.Label(win, text="", fg="blue")
        info_lbl.pack(pady=5)

        frame = tk.Frame(win)
        frame.pack(fill=tk.BOTH, expand=True)
        v_scroll = tk.Scrollbar(frame, orient=tk.VERTICAL)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll = tk.Scrollbar(frame, orient=tk.HORIZONTAL)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas = tk.Canvas(frame, cursor="cross", bg="gray",
                           yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.config(command=canvas.yview)
        h_scroll.config(command=canvas.xview)

        def load_page(page_idx):
            """加载指定页到画布，返回 (scale, offset_x, offset_y)"""
            canvas.update()
            cw = max(canvas.winfo_width(), 100)
            ch = max(canvas.winfo_height(), 100)
            img = self.cleaned_pdf_images[page_idx]
            s = min(cw / img.width, ch / img.height, 1.0)
            dw = int(img.width * s)
            dh = int(img.height * s)
            display = img.resize((dw, dh), Image.Resampling.LANCZOS)
            tk_img = ImageTk.PhotoImage(display)
            canvas.__dict__['_img'] = tk_img
            canvas.delete("all")
            # 画布坐标原点到图片左上角的偏移
            ox = cw // 2 - dw // 2
            oy = ch // 2 - dh // 2
            canvas.create_image(cw // 2, ch // 2, anchor=tk.CENTER, image=tk_img)
            canvas.config(scrollregion=canvas.bbox(tk.ALL))
            return s, ox, oy

        step = 0  # 0 = 设置第1页, 1 = 设置第2页
        rect_id = None
        start_x, start_y = 0, 0
        cur_scale = 1.0
        img_offset_x, img_offset_y = 0, 0

        def refresh_canvas():
            nonlocal rect_id, start_x, start_y, cur_scale, img_offset_x, img_offset_y
            rect_id = None
            start_x, start_y = 0, 0
            cur_scale, img_offset_x, img_offset_y = load_page(step)
            if step == 0:
                info_lbl.config(text="第1步：框选第1页中要丢弃的区域（单数页使用此位置）")
                btn_next.config(text="保存并设置第2页 →", state=tk.NORMAL)
                btn_skip.config(text="跳过此步（不擦除单数页）", state=tk.NORMAL)
            else:
                info_lbl.config(text="第2步：框选第2页中要丢弃的区域（双数页使用此位置）")
                btn_next.config(text="确认并开始OCR", state=tk.NORMAL)
                btn_skip.config(text="跳过（不擦除双数页）", state=tk.NORMAL)

        # 鼠标滚轮
        def on_wheel(event):
            canvas.yview_scroll(-1 if event.delta > 0 else 1, tk.UNITS)
        canvas.bind("<MouseWheel>", on_wheel)

        # 框选交互
        def on_down(event):
            nonlocal start_x, start_y, rect_id
            start_x = canvas.canvasx(event.x)
            start_y = canvas.canvasy(event.y)
            if rect_id is not None:
                canvas.delete(rect_id)
            rect_id = canvas.create_rectangle(
                start_x, start_y, start_x, start_y,
                outline="blue", width=2, dash=(4, 4),
            )

        def on_drag(event):
            if rect_id is not None:
                cur_x = canvas.canvasx(event.x)
                cur_y = canvas.canvasy(event.y)
                canvas.coords(rect_id, start_x, start_y, cur_x, cur_y)

        def on_up(event):
            nonlocal rect_id
            end_x = canvas.canvasx(event.x)
            end_y = canvas.canvasy(event.y)
            x1, x2 = sorted([start_x, end_x])
            y1, y2 = sorted([start_y, end_y])
            if (x2 - x1) > 5 and (y2 - y1) > 5:
                raw = (
                    int((x1 - img_offset_x) / cur_scale),
                    int((y1 - img_offset_y) / cur_scale),
                    int((x2 - img_offset_x) / cur_scale),
                    int((y2 - img_offset_y) / cur_scale),
                )
                self.discard_boxes[step] = raw
            else:
                self.discard_boxes[step] = None
                if rect_id is not None:
                    canvas.delete(rect_id)
                    rect_id = None

        canvas.bind("<ButtonPress-1>", on_down)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_up)

        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=8)

        def on_next():
            nonlocal step
            if step == 0:
                step = 1
                refresh_canvas()
            else:
                win.destroy()
                self._apply_crop_and_ocr()

        def on_skip():
            nonlocal step
            self.discard_boxes[step] = None
            if step == 0:
                step = 1
                refresh_canvas()
            else:
                win.destroy()
                self._apply_crop_and_ocr()

        btn_skip = tk.Button(btn_frame, text="", command=on_skip, width=20)
        btn_skip.pack(side=tk.LEFT, padx=10)
        btn_next = tk.Button(btn_frame, text="", command=on_next, bg="#d1ecf1", width=22)
        btn_next.pack(side=tk.LEFT, padx=10)

        refresh_canvas()

    def _apply_crop_and_ocr(self):
        """按奇偶页分别擦除框选区域，然后执行OCR"""
        self.lbl_pdf.config(text="正在OCR识别...", fg="orange")
        self.root.update()

        pdf_lines = []
        try:
            for i, img in enumerate(self.cleaned_pdf_images):
                box = self.discard_boxes[i % 2]
                if box:
                    draw = ImageDraw.Draw(img)
                    draw.rectangle(box, fill="white")
                text = pytesseract.image_to_string(img, lang='chi_sim', config='--psm 6')
                pdf_lines.append(text)

                self.lbl_pdf.config(text=f"OCR中... 第 {i + 1}/{len(self.cleaned_pdf_images)} 页")
                self.root.update()

            raw_pdf_text = "\n".join(pdf_lines)
            self.pdf_text = self._normalize_text(raw_pdf_text)
            self.lbl_pdf.config(text=f"{self._pdf_name}（OCR完成）", fg="green")
            messagebox.showinfo("完成", "PDF处理完成，现在可以点击「分析差异」进行比对。")
        except Exception as e:
            self.lbl_pdf.config(text=f"{self._pdf_name}（OCR失败）", fg="red")
            messagebox.showerror("OCR错误", f"解析PDF异常:\n{str(e)}")

    def analyze_diff(self):
        if not self.docx_text or not self.pdf_text:
            messagebox.showwarning("提示", "请同时加载Docx和PDF文件后再进行比对。")
            return

        self.txt_diff.delete("1.0", tk.END)
        self.txt_diff.insert(tk.END, "--- DOCX（基准，红色=PDF缺少这些文字）\n", "header")
        self.txt_diff.insert(tk.END, "+++ PDF（绿色=PDF多出这些文字）\n", "header")

        # 去空白、去标点，仅保留纯文字用于比对
        docx_flat = self._keep_only_words(self._strip_whitespace(self.docx_text))
        pdf_flat = self._keep_only_words(self._strip_whitespace(self.pdf_text))

        # 用SequenceMatcher按块比对，连续同状态文字合并展示
        matcher = difflib.SequenceMatcher(None, docx_flat, pdf_flat)
        for op, i1, i2, j1, j2 in matcher.get_opcodes():
            if op == 'equal':
                self.txt_diff.insert(tk.END, docx_flat[i1:i2] + "\n")
            elif op == 'delete':
                self.txt_diff.insert(tk.END, docx_flat[i1:i2] + "\n", "deletion")
            elif op == 'insert':
                self.txt_diff.insert(tk.END, pdf_flat[j1:j2] + "\n", "addition")
            elif op == 'replace':
                self.txt_diff.insert(tk.END, docx_flat[i1:i2] + "\n", "deletion")
                self.txt_diff.insert(tk.END, pdf_flat[j1:j2] + "\n", "addition")

        messagebox.showinfo("分析完成", "Git差异生成完毕。由于OCR存在错字率，请参考变动窗口，对Word执行最终微调。")

    def preview_cleaned_pdf(self):
        """在新窗口中预览去红头后的PDF图像"""
        if not self.cleaned_pdf_images:
            messagebox.showwarning("提示", "请先加载PDF文件。")
            return

        win = tk.Toplevel(self.root)
        win.title("去红头PDF预览")
        img_w, img_h = self.cleaned_pdf_images[0].size
        screen_w = self.root.winfo_screenwidth() - 100
        screen_h = self.root.winfo_screenheight() - 150
        scale = min(screen_w / img_w, screen_h / img_h, 1.0)
        win_w = max(int(img_w * scale), 300)
        win_h = max(int(img_h * scale) + 50, 200)
        win.geometry(f"{win_w}x{win_h}")

        # 翻页控制
        nav_frame = tk.Frame(win, pady=5)
        nav_frame.pack(side=tk.TOP, fill=tk.X)

        # 画布
        frame = tk.Frame(win)
        frame.pack(fill=tk.BOTH, expand=True)
        v_scroll = tk.Scrollbar(frame, orient=tk.VERTICAL)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll = tk.Scrollbar(frame, orient=tk.HORIZONTAL)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas = tk.Canvas(frame, cursor="cross", bg="gray",
                           yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.config(command=canvas.yview)
        h_scroll.config(command=canvas.xview)

        lbl_page = tk.Label(nav_frame, text="")
        lbl_page.pack(side=tk.LEFT, padx=10)
        img_idx = 0

        def show_page():
            # 动态适配画布大小居中显示
            canvas.update()
            cw = max(canvas.winfo_width(), 100)
            ch = max(canvas.winfo_height(), 100)
            img = self.cleaned_pdf_images[img_idx]
            scale = min(cw / img.width, ch / img.height, 1.0)
            w = int(img.width * scale)
            h = int(img.height * scale)
            tk_img = ImageTk.PhotoImage(img.resize((w, h), Image.Resampling.LANCZOS))
            canvas.__dict__['_img'] = tk_img
            canvas.delete("all")
            canvas.create_image(cw // 2, ch // 2, anchor=tk.CENTER, image=tk_img)
            canvas.config(scrollregion=canvas.bbox(tk.ALL))
            lbl_page.config(text=f"第 {img_idx + 1} / {len(self.cleaned_pdf_images)} 页 (缩放 {scale:.0%})")

        def on_wheel(event):
            canvas.yview_scroll(-1 if event.delta > 0 else 1, tk.UNITS)
        canvas.bind("<MouseWheel>", on_wheel)

        def prev():
            nonlocal img_idx
            if img_idx > 0:
                img_idx -= 1
                show_page()

        def next_page():
            nonlocal img_idx
            if img_idx < len(self.cleaned_pdf_images) - 1:
                img_idx += 1
                show_page()

        tk.Button(nav_frame, text="上一页", command=prev).pack(side=tk.LEFT, padx=5)
        tk.Button(nav_frame, text="下一页", command=next_page).pack(side=tk.LEFT, padx=5)
        show_page()

    def save_synced_docx(self):
        """导出依照PDF修改后的Docx：保留格式同步更新"""
        if not self.doc_obj:
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
        if save_path:
            self.doc_obj.save(save_path)
            messagebox.showinfo("成功", "同步更新版（保留格式）文档已导出。")

    def save_dered_docx(self):
        """自动去红头：扫描前部段落，检测红色字体并清除"""
        if not self.doc_obj:
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
        if not save_path:
            return

        # 创建一个用于修改的副本对象
        export_doc = Document(self.docx_path)

        # 自动化去红头策略：
        # 1. 扫描前5个段落。
        # 2. 如果某个段落中存在文字颜色为红色（RGB: >200, <50, <50），
        #    或者包含政府公文红头常见字（如"文件"），则判定为红头，将其内容清空。
        paragraphs_to_check = export_doc.paragraphs[:6]

        for p in paragraphs_to_check:
            is_red_head = False
            # 检查段落中的每个文字运行块（Run）的颜色
            for run in p.runs:
                if run.font.color and run.font.color.rgb:
                    r, g, b = run.font.color.rgb
                    if r > 200 and g < 60 and b < 60:  # 判定为红色系字体
                        is_red_head = True
                        break

            # 如果命中了红头特征，清除该段落文本（保留位置占位，不破坏后续排版结构）
            if is_red_head or "文件" in p.text:
                p.text = ""
                # 连带清除可能遗留的下划线/边框属性
                p.paragraph_format.space_before = 0
                p.paragraph_format.space_after = 0

        try:
            export_doc.save(save_path)
            messagebox.showinfo("成功", "自动去红头版Docx文档已成功导出。")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DocxPdfReviewer(root)
    root.mainloop()
