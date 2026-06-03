"""全局配置 — OCR 引擎实例、共享基类"""
from __future__ import annotations
import tkinter as tk
from typing import Any
from rapidocr_onnxruntime import RapidOCR

ocr_engine = RapidOCR()


class _BaseMixin:
    """所有 Mixin 的共享基类 — 声明跨 Mixin 属性，消除 Pylance 类型警告"""

    # 由 DocxPdfReviewer.__init__ 初始化的实例属性
    root: tk.Tk
    _pairs: list
    _pair_data: dict
    _pair_index: int
    _batch_nav_frame: tk.Frame
    _batch_lbl: tk.Label
    lbl_docx: tk.Label
    lbl_pdf: tk.Label
    lbl_docx_folder: tk.Label
    lbl_pdf_folder: tk.Label
    txt_diff: tk.Text
    docx_path: str | None
    pdf_path: str | None
    _pdf_name: str
    _docx_folder: str | None
    _pdf_folder: str | None
    docx_text: str
    pdf_text: str
    doc_obj: Any | None
    diff_blocks: list
    _docx_paragraphs: list
    _docx_flat_positions: list
    _saved_docx_source: str
    _saved_sel: str | None
    _hover_popup: tk.Toplevel | None
    _hover_block_id: int
    compare_symbols: tk.BooleanVar
    show_info: tk.BooleanVar
    cleaned_pdf_images: list
    discard_boxes: list
    crop_box: Any | None
    _batch_crop_mode: str
    _current_pair_index: int | None

    # 类常量（定义在 DocxPdfReviewer）
    PRIMARY: str
    WARN: str
    SUCCESS: str
    MUTED: str
    BG: str
    TEXT: str
    BORDER: str
    CARD_BG: str

    # 跨 Mixin 方法 stub
    def _alert(self, title: str, message: str, level: str = "info") -> None: ...
    def _set_status(self, text: str, color: str | None = None) -> None: ...
    def _show_crop_dialog(self) -> None: ...
    def _is_red_header(self, paragraph) -> bool: ...
    @staticmethod
    def _normalize_text(text: str) -> str: ...
    @staticmethod
    def _remove_red_pixels(img) -> Any: ...
    @staticmethod
    def _ocr_single(img) -> list: ...
    def _load_docx_path(self, path: str) -> None: ...
    def _load_pdf_path(self, path: str, pair_index: int | None = None) -> None: ...
    def _hide_hover_popup(self) -> None: ...
    def _render_diff(self) -> None: ...
    def _on_diff_motion(self, event) -> None: ...
    def analyze_diff(self, show_warning: bool = True) -> None: ...
    def _on_pair_ocr_done(self, index: int) -> None: ...
    def _save_pair_data(self, index: int) -> None: ...
    def _update_batch_nav(self) -> None: ...
    def _ask_crop_reuse(self) -> None: ...
    def _batch_load_next(self, index: int) -> None: ...
    def _batch_load_next_with_crop(self, index: int) -> None: ...
