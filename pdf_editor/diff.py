"""差异分析 — difflib 比对、Git 风格渲染、悬浮操作窗"""

import tkinter as tk
import difflib
from .config import _BaseMixin

class DiffMixin(_BaseMixin):
    """差异分析 — difflib 比对、Git 风格渲染、悬浮操作窗"""

    def _on_toggle_symbols(self):
        """切换比对符号前，若有已同意的更改则弹窗确认"""
        # Checkbutton 已自动切换了变量值，先恢复
        new_val = self.compare_symbols.get()
        self.compare_symbols.set(not new_val)
        approved = sum(1 for b in self.diff_blocks if b['accepted'] and b['tag'] != 'equal')
        if approved > 0:
            win = tk.Toplevel(self.root)
            win.withdraw()
            win.title("确认切换")
            win.resizable(False, False)
            win.transient(self.root)
            win.grab_set()
            f = tk.Frame(win, padx=20, pady=15, bg="white")
            f.pack(fill=tk.BOTH, expand=True)
            tk.Label(f, text=f"当前有 {approved} 处已同意的更改。",
                     font=("Microsoft YaHei UI", 10), bg="white").pack(pady=(0, 5))
            tk.Label(f, text="切换比对符号将重新分析差异，\n"
                            "已同意的更改可能被重置。确定继续？",
                     font=("Microsoft YaHei UI", 10, "bold"), bg="white").pack(pady=(0, 12))
            btn_frame = tk.Frame(f, bg="white")
            btn_frame.pack()
            tk.Button(btn_frame, text="继续切换",
                      command=lambda: self._do_toggle_symbols(win, new_val),
                      bg=self.PRIMARY, fg="white", relief=tk.FLAT,
                      padx=16, pady=6, cursor="hand2",
                      font=("Microsoft YaHei UI", 10)).pack(side=tk.LEFT, padx=4)
            tk.Button(btn_frame, text="取消",
                      command=win.destroy,
                      bg="#6c757d", fg="white", relief=tk.FLAT,
                      padx=16, pady=6, cursor="hand2",
                      font=("Microsoft YaHei UI", 10)).pack(side=tk.LEFT, padx=4)
            win.bind("<Escape>", lambda _: win.destroy())
            win.update_idletasks()
            pw, ph = self.root.winfo_width(), self.root.winfo_height()
            px, py = self.root.winfo_x(), self.root.winfo_y()
            ww, wh = win.winfo_width(), win.winfo_height()
            win.geometry(f"+{px + (pw - ww) // 2}+{py + (ph - wh) // 2}")
            win.deiconify()
            win.wait_window()
        else:
            self.compare_symbols.set(new_val)
            if not self.docx_text or not self.pdf_text:
                self._alert("提示",
                    "请先加载 DOCX 和 PDF 文件，再进行比对模式切换。\n"
                    "操作步骤：加载 DOCX → 加载 PDF → 裁剪 → OCR → 分析差异",
                    "warn")
                return
            self.analyze_diff(show_warning=False)

    def _do_toggle_symbols(self, win, new_val):
        win.destroy()
        self.compare_symbols.set(new_val)
        self.analyze_diff(show_warning=False)

    def analyze_diff(self, show_warning=True):
        if not self.docx_text or not self.pdf_text:
            if show_warning:
                self._alert("提示", "请同时加载Docx和PDF文件后再进行比对。", "warn")
            return

        if self.compare_symbols.get():
            docx_flat = ''.join(ch for ch in self.docx_text if not ch.isspace())
            pdf_flat = ''.join(ch for ch in self.pdf_text if not ch.isspace())
        else:
            docx_flat = ''.join(
                ch for ch in self.docx_text
                if not ch.isspace() and (
                    '一' <= ch <= '鿿' or '㐀' <= ch <= '䶿'
                    or 'a' <= ch <= 'z' or 'A' <= ch <= 'Z'
                    or '0' <= ch <= '9'
                ))
            pdf_flat = ''.join(
                ch for ch in self.pdf_text
                if not ch.isspace() and (
                    '一' <= ch <= '鿿' or '㐀' <= ch <= '䶿'
                    or 'a' <= ch <= 'z' or 'A' <= ch <= 'Z'
                    or '0' <= ch <= '9'
                ))

        # 构建位置映射表（flat → full_text 位置），用于后续合并
        # 使用与 docx_flat 相同的源文本，确保 save_synced_docx 的二次 diff 一致
        full_text = self.docx_text
        self._saved_docx_source = full_text
        self._docx_flat_positions = []
        for i, ch in enumerate(full_text):
            if self.compare_symbols.get():
                if not ch.isspace():
                    self._docx_flat_positions.append(i)
            else:
                if not ch.isspace() and (
                    '一' <= ch <= '鿿' or '㐀' <= ch <= '䶿'
                    or 'a' <= ch <= 'z' or 'A' <= ch <= 'Z'
                    or '0' <= ch <= '9'
                ):
                    self._docx_flat_positions.append(i)

        # 同时构建原始（未归一化）文本的位置映射，导出时用原始文本写回
        raw_full = '\n'.join(self._docx_paragraphs)
        self._docx_text_raw = raw_full
        self._docx_flat_to_raw = []
        for i, ch in enumerate(raw_full):
            if self.compare_symbols.get():
                if not ch.isspace():
                    self._docx_flat_to_raw.append(i)
            else:
                if not ch.isspace() and (
                    '一' <= ch <= '鿿' or '㐀' <= ch <= '䶿'
                    or 'a' <= ch <= 'z' or 'A' <= ch <= 'Z'
                    or '0' <= ch <= '9'
                ):
                    self._docx_flat_to_raw.append(i)

        # 保留之前已同意的块状态（切换比对符号等场景下避免丢失审核决定）
        prev_accepted = {}
        for b in getattr(self, 'diff_blocks', []) or []:
            if b.get('accepted') and b['tag'] != 'equal':
                prev_accepted[(b['tag'], b['old'], b['new'])] = True

        matcher = difflib.SequenceMatcher(None, docx_flat, pdf_flat)
        self.diff_blocks = []
        block_id = 0
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            old_text = docx_flat[i1:i2]
            new_text = pdf_flat[j1:j2]
            if tag == 'equal':
                self.diff_blocks.append({
                    'tag': tag, 'old': old_text, 'new': '',
                    'accepted': True, 'id': block_id,
                })
            elif tag in ('delete', 'insert', 'replace'):
                was_accepted = prev_accepted.get((tag, old_text, new_text), False)
                self.diff_blocks.append({
                    'tag': tag, 'old': old_text, 'new': new_text,
                    'accepted': was_accepted, 'id': block_id,
                })
            block_id += 1

        self._hide_hover_popup()
        self._render_diff()
        self.txt_diff.bind("<Motion>", self._on_diff_motion)
        self._update_button_states()

        if show_warning:
            self._alert("分析完成",
                "鼠标悬停在更改行上可弹出操作窗口。\n导出时将仅并入已同意的更改。",
                "info")



    def _render_diff(self):
        """根据 self.diff_blocks 渲染差异文本"""
        self.txt_diff.config(state=tk.NORMAL)
        # 保存第一可见行的行号
        try:
            top_line = int(str(self.txt_diff.index("@0,0")).split(".")[0])
        except (ValueError, IndexError):
            top_line = 1
        self.txt_diff.delete("1.0", tk.END)
        mode = "（比对符号）" if self.compare_symbols.get() else "（忽略符号）"
        pair_info = ""
        if self._pairs and self._pair_index >= 0:
            pair_info = f"  |  第 {self._pair_index + 1}/{len(self._pairs)} 对"
        approved = sum(1 for b in self.diff_blocks if b['accepted'] and b['tag'] != 'equal')
        total = sum(1 for b in self.diff_blocks if b['tag'] != 'equal')
        self.txt_diff.insert(tk.END,
            f"--- DOCX（基准）{mode}{pair_info}  |  已同意 {approved}/{total} 处更改\n", "header")
        self.txt_diff.insert(tk.END, "+++ PDF（对比件）\n", "header")

        for b in self.diff_blocks:
            tag = b['tag']
            bid = b['id']

            if tag == 'equal':
                self.txt_diff.insert(tk.END, f"  {b['old']}\n", f"block_{bid}")
            elif tag == 'delete':
                t = "del_on" if b['accepted'] else "del_off"
                self.txt_diff.insert(tk.END, f"- {b['old']}\n",
                                     (f"block_{bid}", t))
            elif tag == 'insert':
                t = "add_on" if b['accepted'] else "add_off"
                self.txt_diff.insert(tk.END, f"+ {b['new']}\n",
                                     (f"block_{bid}", t))
            elif tag == 'replace':
                dt = "del_on" if b['accepted'] else "del_off"
                at = "add_on" if b['accepted'] else "add_off"
                self.txt_diff.insert(tk.END, f"- {b['old']}\n",
                                     (f"block_{bid}", dt))
                self.txt_diff.insert(tk.END, f"+ {b['new']}\n",
                                     (f"block_{bid}", at))

        # 恢复滚动位置
        self.txt_diff.update_idletasks()
        self.txt_diff.see(f"{top_line}.0")

        # 更新提示标签
        if self._diff_hint:
            approved = sum(1 for b in self.diff_blocks if b['accepted'] and b['tag'] != 'equal')
            total = sum(1 for b in self.diff_blocks if b['tag'] != 'equal')
            self._diff_hint.config(
                text=f"鼠标悬停更改行可操作  |  红底=删除  绿底=新增  |  "
                     f"浅色=已忽略  深色=已同意  |  已同意 {approved}/{total} 处更改")



    def _get_block_at(self, x, y):
        """返回鼠标位置所在的 diff 块（仅非 equal 块），无则返回 None"""
        idx = self.txt_diff.index(f"@{x},{y}")
        tags = self.txt_diff.tag_names(idx)
        for tag in tags:
            if tag.startswith("block_"):
                try:
                    bid = int(tag.split("_")[1])
                    for b in self.diff_blocks:
                        if b['id'] == bid and b['tag'] != 'equal':
                            return b
                except (ValueError, IndexError):
                    pass
        return None



    def _on_diff_motion(self, event):
        """鼠标移动：检测悬停的更改块，弹出操作悬浮窗"""
        block = self._get_block_at(event.x, event.y)
        if block is None:
            self._hide_hover_popup()
            return
        if block['id'] == self._hover_block_id:
            return  # 同一块，不刷新

        self._hover_block_id = block['id']
        self._show_hover_popup(block, event)



    def _show_hover_popup(self, block, event):
        """在鼠标附近显示操作悬浮窗"""
        self._hide_hover_popup()

        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.configure(bg="#2d2d2d")

        # 内容
        inner = tk.Frame(popup, bg="#2d2d2d", padx=10, pady=8)
        inner.pack()

        tag_label = {"delete": "删除", "insert": "新增", "replace": "替换"}
        label = tag_label.get(block['tag'], '变更')
        status = "已同意" if block['accepted'] else "已忽略"

        tk.Label(inner, text=f"{label} | {status}",
                 font=("Microsoft YaHei UI", 9, "bold"),
                 bg="#2d2d2d", fg="white").pack(pady=(0, 4))

        # 变更内容预览（截取前30字符）
        if block['tag'] in ('delete', 'replace'):
            preview = block['old'][:30] + ('...' if len(block['old']) > 30 else '')
            tk.Label(inner, text=f"- {preview}",
                     font=("Consolas", 9), bg="#2d2d2d",
                     fg="#ff8888", wraplength=300, justify=tk.LEFT).pack(anchor=tk.W)
        if block['tag'] in ('insert', 'replace'):
            preview = block['new'][:30] + ('...' if len(block['new']) > 30 else '')
            tk.Label(inner, text=f"+ {preview}",
                     font=("Consolas", 9), bg="#2d2d2d",
                     fg="#88ff88", wraplength=300, justify=tk.LEFT).pack(anchor=tk.W)

        # 切换按钮
        btn_text = "改为忽略" if block['accepted'] else "改为同意"
        btn_bg = self.PRIMARY if not block['accepted'] else "#6c757d"
        tk.Button(inner, text=btn_text,
                  command=lambda b=block: self._do_toggle(b),
                  bg=btn_bg, fg="white", relief=tk.FLAT,
                  padx=12, pady=3, cursor="hand2",
                  font=("Microsoft YaHei UI", 9)).pack(pady=(6, 0))

        # 定位：鼠标右下方偏移
        x = self.root.winfo_x() + event.widget.winfo_rootx() - self.root.winfo_rootx() + event.x + 15
        y = self.root.winfo_y() + event.widget.winfo_rooty() - self.root.winfo_rooty() + event.y + 10
        popup.geometry(f"+{x}+{y}")
        self._hover_popup = popup



    def _hide_hover_popup(self):
        if self._hover_popup:
            try:
                self._hover_popup.destroy()
            except tk.TclError:
                pass
            self._hover_popup = None
            self._hover_block_id = -1



    def _on_root_leave(self, event):
        """光标离开窗口时关闭悬浮窗"""
        if not self._hover_popup:
            return
        # 检查光标是否移到了悬浮窗上（悬浮窗是独立 Toplevel）
        try:
            mx = self.root.winfo_pointerx()
            my = self.root.winfo_pointery()
            px = self._hover_popup.winfo_rootx()
            py = self._hover_popup.winfo_rooty()
            pw = self._hover_popup.winfo_width()
            ph = self._hover_popup.winfo_height()
            if pw > 1 and ph > 1 and px <= mx <= px + pw and py <= my <= py + ph:
                return  # 光标在悬浮窗上，不关闭
        except tk.TclError:
            pass
        self._hide_hover_popup()



    def _do_toggle(self, block):
        """切换块状态并刷新"""
        block['accepted'] = not block['accepted']
        self._hide_hover_popup()
        self._render_diff()
        self._update_button_states()

