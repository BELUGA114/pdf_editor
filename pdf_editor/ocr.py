"""OCR 处理 — RapidOCR 识别、裁剪对话框、去红头预览"""

import tkinter as tk
from tkinter import filedialog
import os
import io
import threading
import numpy as np
import fitz
from PIL import Image, ImageTk, ImageDraw
from .config import ocr_engine, _BaseMixin

class OCRMixin(_BaseMixin):
    """OCR 处理 — RapidOCR 识别、裁剪对话框、去红头预览"""
    @staticmethod
    def _ocr_single(img):
        """对单张图片执行OCR"""
        result, _ = ocr_engine(np.array(img))
        lines = []
        if result:
            for _, text, _ in result:
                if text.strip():
                    lines.append(text.strip())
        return lines



    def _apply_crop_and_ocr(self, pair_index: int | None = None):
        """异步执行裁剪+OCR，不阻塞UI。pair_index 非 None 时表示批量模式"""
        self._set_status("正在准备OCR...", "orange")

        # 先同步擦除框选区域（轻量操作）
        for i, img in enumerate(self.cleaned_pdf_images):
            boxes = self.discard_boxes[i % 2]
            if boxes:
                draw = ImageDraw.Draw(img)
                for box in boxes:
                    draw.rectangle(box, fill="white")

        # OCR 放到后台线程，UI 线程用 after 轮询进度
        import threading
        total = len(self.cleaned_pdf_images)
        results: list = []
        progress = {"current": 0}
        error: list = [None]

        def worker():
            try:
                for i, img in enumerate(self.cleaned_pdf_images):
                    lines = self._ocr_single(img)
                    results.append((i, lines))
                    progress["current"] = i + 1
            except Exception as e:
                error[0] = e

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

        def poll():
            cur = progress["current"]
            suffix = f" (第{pair_index + 1}/{len(self._pairs)}对)" if pair_index is not None and self._pairs else ""
            self._set_status(f"OCR中... 第 {cur}/{total} 页{suffix}", "orange")
            if thread.is_alive():
                self.root.after(150, poll)
            elif error[0]:
                self.lbl_pdf.config(text=f"{self._pdf_name}（OCR失败）", fg="red")
                self._alert("OCR错误", f"解析PDF异常:\n{str(error[0])}", "error")
                self._set_status("")
            else:
                results.sort(key=lambda x: x[0])
                pdf_lines = []
                for _, lines in results:
                    pdf_lines.extend(lines)
                raw_pdf_text = "\n".join(pdf_lines)
                self.pdf_text = self._normalize_text(raw_pdf_text)
                self.lbl_pdf.config(text=f"{self._pdf_name}（OCR完成）", fg="green")
                self._set_status("")
                if pair_index is not None:
                    self._on_pair_ocr_done(pair_index)
                else:
                    self._alert("完成", "PDF处理完成，现在可以点击「分析差异」进行比对。", "info")

        self.root.after(100, poll)



    def _show_crop_dialog(self):
        """两步裁剪：第1页（单数页基准）→ 第2页（双数页基准），每页可框选多个区域"""
        if not self.cleaned_pdf_images:
            return

        self.discard_boxes: list = [[], []]  # 每页一个列表，存放多个 (x1,y1,x2,y2)
        win = tk.Toplevel(self.root)
        win.title("裁剪PDF（框选要丢弃的区域，可多选）")
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
            ox = cw // 2 - dw // 2
            oy = ch // 2 - dh // 2
            canvas.create_image(cw // 2, ch // 2, anchor=tk.CENTER, image=tk_img)
            canvas.config(scrollregion=canvas.bbox(tk.ALL))
            # 重绘已保存的框选区域
            for box in self.discard_boxes[page_idx]:
                x1, y1, x2, y2 = box
                canvas.create_rectangle(
                    ox + x1 * s, oy + y1 * s, ox + x2 * s, oy + y2 * s,
                    outline="blue", width=2, dash=(4, 4),
                )
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
            count = len(self.discard_boxes[step])
            suffix = f"（已选 {count} 个区域）" if count else ""
            if step == 0:
                info_lbl.config(text=f"第1步：框选第1页中要丢弃的区域{suffix}（单数页使用此位置）")
                btn_next.config(text="保存并设置第2页 →", state=tk.NORMAL)
                btn_skip.config(text="跳过此步（不擦除单数页）", state=tk.NORMAL)
                btn_undo.config(state=tk.NORMAL if count else tk.DISABLED)
            else:
                info_lbl.config(text=f"第2步：框选第2页中要丢弃的区域{suffix}（双数页使用此位置）")
                btn_next.config(text="确认并开始OCR", state=tk.NORMAL)
                btn_skip.config(text="跳过（不擦除双数页）", state=tk.NORMAL)
                btn_undo.config(state=tk.NORMAL if count else tk.DISABLED)

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
                outline="red", width=2, dash=(4, 4),
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
                self.discard_boxes[step].append(raw)
                # 把当前矩形从临时色改为持久色
                if rect_id is not None:
                    canvas.itemconfig(rect_id, outline="blue")
                rect_id = None
                refresh_canvas()
            else:
                if rect_id is not None:
                    canvas.delete(rect_id)
                    rect_id = None

        canvas.bind("<ButtonPress-1>", on_down)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_up)

        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=8)

        def on_undo():
            if self.discard_boxes[step]:
                self.discard_boxes[step].pop()
                refresh_canvas()

        def on_next():
            nonlocal step
            if step == 0:
                step = 1
                refresh_canvas()
            else:
                win.destroy()
                pi = getattr(self, '_current_pair_index', None)
                self._apply_crop_and_ocr(pair_index=pi)

        def on_skip():
            nonlocal step
            self.discard_boxes[step] = []
            if step == 0:
                step = 1
                refresh_canvas()
            else:
                win.destroy()
                pi = getattr(self, '_current_pair_index', None)
                self._apply_crop_and_ocr(pair_index=pi)

        btn_undo = tk.Button(btn_frame, text="撤销上一个区域", command=on_undo, width=18,
                             state=tk.DISABLED)
        btn_undo.pack(side=tk.LEFT, padx=10)
        btn_skip = tk.Button(btn_frame, text="", command=on_skip, width=20)
        btn_skip.pack(side=tk.LEFT, padx=10)
        btn_next = tk.Button(btn_frame, text="", command=on_next, bg="#d1ecf1", width=22)
        btn_next.pack(side=tk.LEFT, padx=10)

        refresh_canvas()

    def preview_cleaned_pdf(self):
        """在新窗口中预览去红头后的PDF图像"""
        if not self.cleaned_pdf_images:
            self._alert("提示", "请先加载PDF文件。", "warn")
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

