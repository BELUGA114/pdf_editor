"""Windows 拖放支持 — 注册窗口接受文件拖放消息"""
import ctypes
from ctypes import wintypes


def _enable_dnd(hwnd: int, callback):
    """为 Windows 窗口启用文件拖放，拖入文件时调用 callback(paths)"""
    user32 = ctypes.windll.user32
    shell32 = ctypes.windll.shell32

    # 显式声明 API 签名，避免 64 位指针截断
    user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.GetWindowLongPtrW.restype = wintypes.LPARAM
    user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.LPARAM]
    user32.SetWindowLongPtrW.restype = wintypes.LPARAM

    shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
    shell32.DragAcceptFiles(hwnd, True)
    # 恢复默认 argtypes，避免影响 CallWindowProcW 等后续调用
    del shell32.DragAcceptFiles.argtypes

    GWLP_WNDPROC = -4
    WM_DROPFILES = 0x0233

    WNDPROC = ctypes.WINFUNCTYPE(
        wintypes.LPARAM, wintypes.HWND, wintypes.UINT,
        wintypes.WPARAM, wintypes.LPARAM,
    )
    original = wintypes.LPARAM(user32.GetWindowLongPtrW(hwnd, GWLP_WNDPROC))

    @WNDPROC
    def new_wndproc(hwnd, msg, wparam, lparam):
        if msg == WM_DROPFILES:
            shell32.DragQueryFileW.argtypes = [wintypes.HANDLE, wintypes.UINT,
                                                wintypes.LPWSTR, wintypes.UINT]
            shell32.DragQueryFileW.restype = wintypes.UINT
            shell32.DragFinish.argtypes = [wintypes.HANDLE]
            try:
                hdrop = wintypes.HANDLE(wparam)
                count = shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
                paths = []
                for i in range(count):
                    buf = ctypes.create_unicode_buffer(260)
                    shell32.DragQueryFileW(hdrop, i, buf, 260)
                    paths.append(buf.value)
                shell32.DragFinish(hdrop)
                callback(paths)
            finally:
                del shell32.DragQueryFileW.argtypes, shell32.DragQueryFileW.restype
                del shell32.DragFinish.argtypes
            return 0
        return user32.CallWindowProcW(original,
            wintypes.HWND(hwnd), wintypes.UINT(msg),
            wintypes.WPARAM(wparam), wintypes.LPARAM(lparam))

    user32.SetWindowLongPtrW(hwnd, GWLP_WNDPROC,
                              ctypes.cast(new_wndproc, ctypes.c_void_p).value)
    return new_wndproc
