"""Docx-PDF 差异比对工具 — 主应用类"""
import tkinter as tk
from tkinter import filedialog
import os

from .dnd import _enable_dnd
from .loader import LoaderMixin
from .ocr import OCRMixin
from .batch import BatchMixin
from .diff import DiffMixin
from .export import ExportMixin


class DocxPdfReviewer(LoaderMixin, OCRMixin, BatchMixin, DiffMixin, ExportMixin):
    # 颜色常量
    BG = "#f0f2f5"         # 页面背景
    CARD_BG = "#ffffff"    # 卡片背景
    PRIMARY = "#4a6eb5"    # 主色调
    SUCCESS = "#52c41a"    # 成功/导出
    WARN = "#faad14"       # 警告
    BORDER = "#d9d9d9"     # 边框
    TEXT = "#333333"       # 正文
    MUTED = "#999999"      # 弱化文字



    def __init__(self, root):
        self.root = root
        self.root.title("Docx-PDF 差异比对工具")
        self.root.geometry("1200x800")
        self.root.configure(bg=self.BG)
        # 全局字体
        self.root.option_add("*Font", ("Microsoft YaHei UI", 10))

        self.docx_path = None
        self.pdf_path = None
        self.doc_obj = None

        self.docx_text = ""
        self.pdf_text = ""
        self.diff_results = []
        self.cleaned_pdf_images = []
        self._saved_sel = None
        self.compare_symbols = tk.BooleanVar(value=True)
        self.show_info = tk.BooleanVar(value=True)
        self._docx_folder = None
        self._pdf_folder = None
        self._pairs = []           # 批量比对: [(docx_path, pdf_path), ...]
        self._pair_index = -1
        self._pair_data = {}       # {index: {docx_text, pdf_text, doc_obj, cleaned, ...}}
        self.diff_blocks = []      # [{tag, old, new, accepted, block_id}, ...]
        self._docx_paragraphs = [] # 原始段落文本（用于合并回写）
        self._docx_flat_positions = []  # flat → full_text 位置映射
        self._saved_docx_source = ''    # 分析时的源文本（保证回写一致性）
        self._hover_popup = None   # 悬浮窗引用
        self._hover_block_id = -1  # 当前悬浮的块ID

        self._build_ui()

        # 光标离开窗口时关闭悬浮窗
        self.root.bind("<Leave>", self._on_root_leave)

        # 注册拖放（需在窗口实现后）
        self.root.update_idletasks()
        self._dnd_proc = _enable_dnd(self.root.winfo_id(), self._on_drop)



    def _build_ui(self):
        # ---- 顶部标题 ----
        header = tk.Frame(self.root, bg="#1a1a2e", padx=20, pady=12)
        header.pack(fill=tk.X)
        tk.Label(header, text="Docx ↔ PDF 差异比对",
                 font=("Microsoft YaHei UI", 16, "bold"),
                 bg="#1a1a2e", fg="white").pack(side=tk.LEFT)
        tk.Label(header, text="OCR 驱动 · Git 风格展示",
                 font=("Microsoft YaHei UI", 9),
                 bg="#1a1a2e", fg="#a0a0c0").pack(side=tk.LEFT, padx=12)

        # ---- 拖放区域 2x2 网格 ----
        grid_frame = tk.Frame(self.root, bg=self.BG, padx=20, pady=10)
        grid_frame.pack(fill=tk.X)
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.columnconfigure(1, weight=1)

        docx_card, self.lbl_docx = self._make_drop_zone(
            grid_frame, "拖入 DOCX", "原始 Word 文档（基准）", "📄",
            self.load_docx)
        docx_card.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))

        pdf_card, self.lbl_pdf = self._make_drop_zone(
            grid_frame, "拖入 PDF", "扫描版 PDF（对比件）", "📑",
            self.load_pdf)
        pdf_card.grid(row=0, column=1, sticky="nsew", padx=(4, 0), pady=(0, 4))

        docx_folder_card, self.lbl_docx_folder = self._make_drop_zone(
            grid_frame, "选择 DOCX 文件夹", "批量 Word 文档目录", "📁",
            self._select_docx_folder)
        docx_folder_card.grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(4, 0))

        pdf_folder_card, self.lbl_pdf_folder = self._make_drop_zone(
            grid_frame, "选择 PDF 文件夹", "批量 PDF 文件目录", "📂",
            self._select_pdf_folder)
        pdf_folder_card.grid(row=1, column=1, sticky="nsew", padx=(4, 0), pady=(4, 0))

        # ---- 操作栏 ----
        action_frame = tk.Frame(self.root, bg=self.BG, padx=20, pady=10)
        action_frame.pack(fill=tk.X)

        self._make_btn(action_frame, "分析差异 (Git 模式)",
                       self.analyze_diff).pack(side=tk.LEFT, padx=(0, 8))
        tk.Checkbutton(action_frame, text="比对符号", variable=self.compare_symbols,
                       command=lambda: self.analyze_diff(show_warning=False),
                       bg=self.BG, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT, padx=8)
        tk.Checkbutton(action_frame, text="提示弹窗", variable=self.show_info,
                       bg=self.BG, font=("Microsoft YaHei UI", 9)).pack(side=tk.LEFT)
        self._make_btn(action_frame, "预览去红头 PDF", self.preview_cleaned_pdf,
                       bg="#e8a838").pack(side=tk.LEFT, padx=8)

        self._status_lbl = tk.Label(action_frame, text="",
                                    font=("Microsoft YaHei UI", 9),
                                    bg=self.BG, fg=self.MUTED)
        self._status_lbl.pack(side=tk.LEFT, padx=8)

        # 导出按钮（右侧）
        export_frame = tk.Frame(action_frame, bg=self.BG)
        export_frame.pack(side=tk.RIGHT)
        self._make_btn(export_frame, "导出：同步更新版 Docx", self.save_synced_docx,
                       bg=self.SUCCESS, font=("Microsoft YaHei UI", 9)).pack(
            side=tk.LEFT, padx=4)
        self._make_btn(export_frame, "导出：自动去红头版 Docx", self.save_dered_docx,
                       bg="#e8a838", font=("Microsoft YaHei UI", 9)).pack(
            side=tk.LEFT, padx=4)

        # ---- 批量导航栏（初始隐藏） ----
        self._batch_nav_frame = tk.Frame(self.root, bg=self.BG)
        # 不 pack，等批量模式激活时再显示
        self._make_btn(self._batch_nav_frame, "← 上一对", self._nav_prev,
                       bg="#6c757d", font=("Microsoft YaHei UI", 9)).pack(
            side=tk.LEFT, padx=(0, 4))
        self._batch_lbl = tk.Label(self._batch_nav_frame, text="第 1/1 对",
                                   font=("Microsoft YaHei UI", 9, "bold"),
                                   bg=self.BG, fg=self.TEXT)
        self._batch_lbl.pack(side=tk.LEFT, padx=8)
        self._make_btn(self._batch_nav_frame, "下一对 →", self._nav_next,
                       bg="#6c757d", font=("Microsoft YaHei UI", 9)).pack(
            side=tk.LEFT, padx=4)

        # ---- 差异展示区 ----
        mid_frame = tk.Frame(self.root, padx=20, pady=15, bg=self.BG)
        mid_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(mid_frame,
                 text="鼠标悬停更改行可操作  |  红底=删除  绿底=新增  |  浅色=已忽略  深色=已同意",
                 font=("Microsoft YaHei UI", 9),
                 bg=self.BG, fg=self.MUTED).pack(anchor=tk.W, pady=(0, 6))

        text_frame = tk.Frame(mid_frame, bg=self.CARD_BG,
                              highlightbackground=self.BORDER, highlightthickness=1)
        text_frame.pack(fill=tk.BOTH, expand=True)

        self.txt_diff = tk.Text(text_frame, font=("Consolas", 11), wrap=tk.WORD,
                                exportselection=True,
                                selectbackground=self.PRIMARY, selectforeground="white",
                                inactiveselectbackground="#7a9ec5",
                                bg=self.CARD_BG, fg=self.TEXT, bd=0,
                                padx=12, pady=10)
        self.txt_diff.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll = tk.Scrollbar(text_frame, command=self.txt_diff.yview, bg=self.BG)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_diff.config(yscrollcommand=scroll.set)

        self.txt_diff.tag_config("del_on", background="#ffd4d4", foreground="#b30000",
                                 font=("Consolas", 11, "bold"))
        self.txt_diff.tag_config("add_on", background="#c8f0c8", foreground="#008000",
                                 font=("Consolas", 11, "bold"))
        self.txt_diff.tag_config("del_off", background="#fff5f5", foreground="#ccaaaa")
        self.txt_diff.tag_config("add_off", background="#f5fff5", foreground="#aaccaa")
        self.txt_diff.tag_config("header", foreground="#6a737d",
                                 font=("Consolas", 10, "italic"))
        self.txt_diff.tag_raise("sel")

        self._setup_copy_menu(self.txt_diff)



    def _on_drop(self, paths):
        for path in paths:
            if os.path.isdir(path):
                self._on_drop_folder(path)
            else:
                ext = os.path.splitext(path)[1].lower()
                if ext == ".docx":
                    self._load_docx_path(path)
                elif ext == ".pdf":
                    self._load_pdf_path(path)



    def _on_drop_folder(self, path):
        """拖入文件夹时自动判断类型（按文件数占比）"""
        docx_cnt = 0
        pdf_cnt = 0
        try:
            for f in os.listdir(path):
                ext = os.path.splitext(f)[1].lower()
                if ext in {'.docx', '.doc'}:
                    docx_cnt += 1
                elif ext == '.pdf':
                    pdf_cnt += 1
        except OSError:
            return

        if docx_cnt == 0 and pdf_cnt == 0:
            return

        if docx_cnt >= pdf_cnt:
            self._docx_folder = path
            self.lbl_docx_folder.config(text=os.path.basename(path), fg=self.PRIMARY)
        if pdf_cnt >= docx_cnt:
            self._pdf_folder = path
            self.lbl_pdf_folder.config(text=os.path.basename(path), fg=self.PRIMARY)

        self._try_folder_match()



    def _set_status(self, text: str, color: str | None = None):
        """更新操作栏状态标签"""
        self._status_lbl.config(text=text, fg=color or self.MUTED)
        self.root.update_idletasks()



    def _alert(self, title: str, message: str, level: str = "info"):
        """无提示音的消息弹窗（替代 messagebox）"""
        if level == "info" and not self.show_info.get():
            return
        colors = {"info": "#d1ecf1", "warn": "#fff3cd", "error": "#f8d7da"}
        win = tk.Toplevel(self.root)
        win.withdraw()
        win.title(title)
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        bg = colors.get(level, "#d1ecf1")
        f = tk.Frame(win, padx=20, pady=15, bg=bg)
        f.pack(fill=tk.BOTH, expand=True)
        tk.Label(f, text=message, bg=bg, wraplength=400, justify=tk.LEFT,
                 font=("Microsoft YaHei UI", 10)).pack(pady=(0, 10))
        btn = tk.Button(f, text="确定", command=win.destroy, width=10,
                        bg=self.PRIMARY, fg="white", relief=tk.FLAT,
                        padx=20, pady=4, cursor="hand2")
        btn.pack()
        btn.focus_set()
        win.bind("<Return>", lambda _: win.destroy())
        win.bind("<Escape>", lambda _: win.destroy())
        win.update_idletasks()
        pw, ph = self.root.winfo_width(), self.root.winfo_height()
        px, py = self.root.winfo_x(), self.root.winfo_y()
        ww, wh = win.winfo_width(), win.winfo_height()
        x = px + (pw - ww) // 2
        y = py + (ph - wh) // 2
        win.geometry(f"+{x}+{y}")
        win.deiconify()
        win.wait_window()



    def _make_drop_zone(self, parent, title: str, subtitle: str, icon: str, on_click):
        """创建一个可点击/拖放的文件卡片区域（水平布局：图标 | 文字）"""
        card = tk.Frame(parent, bg=self.CARD_BG, highlightbackground=self.BORDER,
                        highlightthickness=1, cursor="hand2")
        inner = tk.Frame(card, bg=self.CARD_BG, padx=10, pady=6, cursor="hand2")
        inner.pack(fill=tk.BOTH, expand=True)

        def on_enter(_):
            card.config(highlightbackground=self.PRIMARY, highlightthickness=2)
        def on_leave(_):
            card.config(highlightbackground=self.BORDER, highlightthickness=1)

        widgets: list = [card, inner]

        # 左侧图标
        icon_lbl = tk.Label(inner, text=icon, font=("Segoe UI", 16), bg=self.CARD_BG,
                            cursor="hand2")
        icon_lbl.pack(side=tk.LEFT, padx=(0, 8))
        widgets.append(icon_lbl)

        # 右侧文字区域
        text_col = tk.Frame(inner, bg=self.CARD_BG, cursor="hand2")
        text_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        widgets.append(text_col)

        lbl_title = tk.Label(text_col, text=title, font=("Microsoft YaHei UI", 9, "bold"),
                             bg=self.CARD_BG, fg=self.TEXT, cursor="hand2",
                             anchor=tk.W)
        lbl_title.pack(fill=tk.X)
        widgets.append(lbl_title)

        sub_lbl = tk.Label(text_col, text=subtitle, font=("Microsoft YaHei UI", 7),
                           bg=self.CARD_BG, fg=self.MUTED, cursor="hand2",
                           anchor=tk.W)
        sub_lbl.pack(fill=tk.X)
        widgets.append(sub_lbl)

        status_lbl = tk.Label(text_col, text="未加载", font=("Microsoft YaHei UI", 7),
                              bg=self.CARD_BG, fg=self.MUTED, cursor="hand2",
                              anchor=tk.W)
        status_lbl.pack(fill=tk.X)
        widgets.append(status_lbl)

        for w in widgets:
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)
            w.bind("<Button-1>", lambda _: on_click())
        return card, status_lbl



    def _make_btn(self, parent, text: str, command, bg=None, **kw):
        """统一样式的按钮"""
        b = tk.Button(parent, text=text, command=command,
                      bg=bg or self.PRIMARY, fg="white",
                      relief=tk.FLAT, padx=14, pady=6,
                      cursor="hand2", **kw)
        return b



    def _setup_copy_menu(self, widget):
        """为Text组件添加右键复制菜单"""
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="复制", command=lambda: self._do_copy(widget))

        def on_right_click(e):
            try:
                self._saved_sel = widget.selection_get()
            except tk.TclError:
                self._saved_sel = None
            menu.post(e.x_root, e.y_root)

        widget.bind("<Button-3>", on_right_click)
        widget.bind("<Control-c>", lambda _: self._do_copy(widget))



    def _do_copy(self, widget):
        try:
            sel = getattr(self, '_saved_sel', None) or widget.selection_get()
            widget.clipboard_clear()
            widget.clipboard_append(sel)
        except tk.TclError:
            pass


def main():
    """应用入口"""
    root = tk.Tk()
    app = DocxPdfReviewer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
