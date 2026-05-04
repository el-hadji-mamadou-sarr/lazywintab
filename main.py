"""
LazyWinTab - Custom vertical Alt-Tab replacement for Windows.

Usage:
  - Press Alt+Tab to open the switcher.
  - While holding Alt, press Tab to cycle down, Shift+Tab to cycle up.
  - Release Alt to switch to the selected window.
  - Press Escape to dismiss without switching.
  - Click on any entry to switch immediately.
"""

import sys
import os
import ctypes
import ctypes.wintypes as wintypes
import logging
import winreg
from ctypes import POINTER, WINFUNCTYPE, byref, sizeof, c_int, windll

from PyQt6.QtCore import Qt, QSize, QTimer, QRect
from PyQt6.QtGui import QImage, QPixmap, QIcon, QKeyEvent, QFont, QColor, QPainter
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QListWidget, QListWidgetItem, QAbstractItemView, QSystemTrayIcon, QMenu,
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("lazywintab")

# ---------------------------------------------------------------------------
# Win32 constants & helpers
# ---------------------------------------------------------------------------

GWL_EXSTYLE = -20
GW_OWNER = 4
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000
WS_EX_NOACTIVATE = 0x08000000
GA_ROOTOWNER = 3
ICON_SMALL = 0
ICON_BIG = 1
ICON_SMALL2 = 2
WM_GETICON = 0x007F
SMTO_ABORTIFHUNG = 0x0002
GCL_HICON = -14
GCL_HICONSM = -34
SW_RESTORE = 9
SW_SHOW = 5
MOD_ALT = 0x0001
MOD_NOREPEAT = 0x4000
WM_HOTKEY = 0x0312
HOTKEY_ID = 1

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
shell32 = ctypes.windll.shell32

user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, c_int]
user32.GetWindowTextW.restype = c_int
user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextLengthW.restype = c_int
user32.IsWindowVisible.argtypes = [wintypes.HWND]
user32.IsWindowVisible.restype = wintypes.BOOL
user32.GetWindowLongW.argtypes = [wintypes.HWND, c_int]
user32.GetWindowLongW.restype = ctypes.c_long
user32.GetWindow.argtypes = [wintypes.HWND, wintypes.UINT]
user32.GetWindow.restype = wintypes.HWND
user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
user32.GetAncestor.restype = wintypes.HWND
user32.GetLastActivePopup.argtypes = [wintypes.HWND]
user32.GetLastActivePopup.restype = wintypes.HWND
user32.IsWindow.argtypes = [wintypes.HWND]
user32.IsWindow.restype = wintypes.BOOL
user32.SetForegroundWindow.argtypes = [wintypes.HWND]
user32.SetForegroundWindow.restype = wintypes.BOOL
user32.ShowWindow.argtypes = [wintypes.HWND, c_int]
user32.ShowWindow.restype = wintypes.BOOL
user32.IsIconic.argtypes = [wintypes.HWND]
user32.IsIconic.restype = wintypes.BOOL
user32.SendMessageTimeoutW.argtypes = [
    wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM,
    wintypes.UINT, wintypes.UINT, POINTER(ctypes.c_void_p),
]
user32.SendMessageTimeoutW.restype = ctypes.c_long
user32.GetClassLongPtrW.argtypes = [wintypes.HWND, c_int]
user32.GetClassLongPtrW.restype = ctypes.POINTER(ctypes.c_ulong)
user32.DestroyIcon.argtypes = [wintypes.HICON]
user32.DestroyIcon.restype = wintypes.BOOL
user32.GetAsyncKeyState.argtypes = [c_int]
user32.GetAsyncKeyState.restype = ctypes.c_short
user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.PostMessageW.restype = wintypes.BOOL
user32.SetWindowsHookExW.argtypes = [c_int, ctypes.c_void_p, wintypes.HINSTANCE, wintypes.DWORD]
user32.SetWindowsHookExW.restype = ctypes.c_void_p
user32.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]
user32.UnhookWindowsHookEx.restype = wintypes.BOOL
user32.CallNextHookEx.argtypes = [ctypes.c_void_p, c_int, wintypes.WPARAM, wintypes.LPARAM]
user32.CallNextHookEx.restype = ctypes.c_long

WH_KEYBOARD_LL = 13
WH_MOUSE_LL = 14
WM_KEYDOWN = 0x0100
WM_SYSKEYDOWN = 0x0104
WM_MOUSEWHEEL = 0x020A


class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", wintypes.DWORD),
        ("scanCode", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


wintypes.KBDLLHOOKSTRUCT = KBDLLHOOKSTRUCT


class MSLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("pt", wintypes.POINT),
        ("mouseData", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]

ENUMWINDOWSPROC = WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

VK_MENU = 0x12  # Alt key
VK_TAB = 0x09
VK_LMENU = 0xA4
VK_RMENU = 0xA5


def _get_window_text(hwnd: int) -> str:
    length = user32.GetWindowTextLengthW(hwnd)
    if length == 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value


def _is_alt_tab_window(hwnd: int) -> bool:
    """Return True if the window would normally appear in Alt-Tab."""
    if not user32.IsWindowVisible(hwnd):
        return False
    # Walk up owner chain
    root = user32.GetAncestor(hwnd, GA_ROOTOWNER)
    if _get_last_visible_popup(root) != hwnd:
        return False
    ex_style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    if ex_style & WS_EX_TOOLWINDOW:
        return False
    if ex_style & WS_EX_NOACTIVATE:
        return False
    title = _get_window_text(hwnd)
    if not title:
        return False
    return True


def _get_last_visible_popup(hwnd: int) -> int:
    popup = user32.GetLastActivePopup(hwnd)
    if user32.IsWindowVisible(popup):
        return popup
    return hwnd


kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
kernel32.CloseHandle.restype = wintypes.BOOL
_psapi = ctypes.WinDLL("psapi")
_psapi.GetModuleBaseNameW.argtypes = [wintypes.HANDLE, wintypes.HMODULE, wintypes.LPWSTR, wintypes.DWORD]
_psapi.GetModuleBaseNameW.restype = wintypes.DWORD
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
PROCESS_VM_READ = 0x0010

user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD


def _get_process_name(hwnd: int) -> str:
    pid = wintypes.DWORD(0)
    user32.GetWindowThreadProcessId(hwnd, byref(pid))
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_VM_READ, False, pid.value)
    if not handle:
        return ""
    buf = ctypes.create_unicode_buffer(260)
    _psapi.GetModuleBaseNameW(handle, None, buf, 260)
    kernel32.CloseHandle(handle)
    # Strip .exe suffix
    name = buf.value
    if name.lower().endswith(".exe"):
        name = name[:-4]
    return name


def enumerate_windows() -> list[tuple[int, str, str]]:
    """Return (hwnd, title, process_name) for all Alt-Tab eligible windows, in MRU order."""
    results: list[tuple[int, str, str]] = []

    @ENUMWINDOWSPROC
    def callback(hwnd, _lparam):
        if _is_alt_tab_window(hwnd):
            title = _get_window_text(hwnd)
            proc = _get_process_name(hwnd)
            results.append((hwnd, title, proc))
        return True

    user32.EnumWindows(callback, 0)
    return results


def group_windows(windows: list[tuple[int, str, str]]) -> list[tuple[int, str, str]]:
    """Group windows by process name, preserving MRU order of the first window per group."""
    # EnumWindows returns Z-order (MRU first)
    seen: dict[str, list] = {}
    for entry in windows:
        key = entry[2].lower()  # process name
        seen.setdefault(key, []).append(entry)

    # Rebuild: for each group in order of first appearance, emit all its windows
    ordered: list[tuple[int, str, str]] = []
    emitted: set[str] = set()
    for entry in windows:
        key = entry[2].lower()
        if key not in emitted:
            emitted.add(key)
            ordered.extend(seen[key])
    return ordered


def get_window_icon(hwnd: int, size: int = 32) -> QPixmap | None:
    """Try to extract the window icon for the given hwnd."""
    hicon = ctypes.c_void_p(0)

    # Method 1: WM_GETICON
    user32.SendMessageTimeoutW(
        hwnd, WM_GETICON, ICON_BIG, 0, SMTO_ABORTIFHUNG, 100, byref(hicon)
    )
    if not hicon.value:
        user32.SendMessageTimeoutW(
            hwnd, WM_GETICON, ICON_SMALL2, 0, SMTO_ABORTIFHUNG, 100, byref(hicon)
        )
    if not hicon.value:
        user32.SendMessageTimeoutW(
            hwnd, WM_GETICON, ICON_SMALL, 0, SMTO_ABORTIFHUNG, 100, byref(hicon)
        )

    # Method 2: GetClassLongPtr
    if not hicon.value:
        try:
            val = user32.GetClassLongPtrW(hwnd, GCL_HICON)
            if val:
                hicon.value = ctypes.cast(val, ctypes.c_void_p).value
        except Exception:
            pass
    if not hicon.value:
        try:
            val = user32.GetClassLongPtrW(hwnd, GCL_HICONSM)
            if val:
                hicon.value = ctypes.cast(val, ctypes.c_void_p).value
        except Exception:
            pass

    if not hicon.value:
        return None

    return _hicon_to_pixmap(int(hicon.value), size)


def _hicon_to_pixmap(hicon: int, size: int) -> QPixmap | None:
    """Convert a HICON to a QPixmap using GetIconInfo + DIB."""
    try:
        class ICONINFO(ctypes.Structure):
            _fields_ = [
                ("fIcon", wintypes.BOOL),
                ("xHotspot", wintypes.DWORD),
                ("yHotspot", wintypes.DWORD),
                ("hbmMask", wintypes.HBITMAP),
                ("hbmColor", wintypes.HBITMAP),
            ]

        class BITMAPINFOHEADER(ctypes.Structure):
            _fields_ = [
                ("biSize", wintypes.DWORD),
                ("biWidth", wintypes.LONG),
                ("biHeight", wintypes.LONG),
                ("biPlanes", wintypes.WORD),
                ("biBitCount", wintypes.WORD),
                ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD),
                ("biXPelsPerMeter", wintypes.LONG),
                ("biYPelsPerMeter", wintypes.LONG),
                ("biClrUsed", wintypes.DWORD),
                ("biClrImportant", wintypes.DWORD),
            ]

        ii = ICONINFO()
        if not user32.GetIconInfo(hicon, byref(ii)):
            return None

        gdi32 = ctypes.windll.gdi32
        hdc = user32.GetDC(0)

        bmi = BITMAPINFOHEADER()
        bmi.biSize = sizeof(BITMAPINFOHEADER)
        bmi.biWidth = size
        bmi.biHeight = -size  # top-down
        bmi.biPlanes = 1
        bmi.biBitCount = 32
        bmi.biCompression = 0  # BI_RGB

        buf = (ctypes.c_uint8 * (size * size * 4))()
        hbm = ii.hbmColor if ii.hbmColor else ii.hbmMask

        gdi32.GetDIBits(hdc, hbm, 0, size, buf, byref(bmi), 0)
        user32.ReleaseDC(0, hdc)

        if ii.hbmColor:
            gdi32.DeleteObject(ii.hbmColor)
        if ii.hbmMask:
            gdi32.DeleteObject(ii.hbmMask)

        img = QImage(bytes(buf), size, size, QImage.Format.Format_ARGB32)
        return QPixmap.fromImage(img)
    except Exception:
        return None


def switch_to_window(hwnd: int) -> None:
    """Bring the given window to the foreground."""
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, SW_RESTORE)
    user32.SetForegroundWindow(hwnd)
    log.info("Switched to hwnd=%s", hwnd)


# ---------------------------------------------------------------------------
# Switcher UI
# ---------------------------------------------------------------------------

STYLESHEET = """
QWidget#SwitcherBackground {
    background-color: rgba(22, 22, 28, 252);
    border: 1px solid rgba(255, 255, 255, 18);
    border-radius: 16px;
}

QListWidget {
    background: transparent;
    border: none;
    outline: none;
    padding: 6px 8px;
}

QListWidget::item {
    background: transparent;
    border-radius: 10px;
    margin: 1px 0px;
    color: #e8e8e8;
}

QListWidget::item:selected {
    background-color: rgba(120, 80, 220, 160);
}

QListWidget::item:hover:!selected {
    background-color: rgba(255, 255, 255, 12);
}

QLabel#Header {
    color: rgba(160, 160, 170, 180);
    font-size: 10px;
    letter-spacing: 1px;
    padding: 10px 16px 4px 16px;
}
"""

ITEM_HEIGHT = 62
MAX_VISIBLE_ITEMS = 10
WINDOW_WIDTH = 500

import re as _re

def _format_title(raw_title: str, proc: str) -> tuple[str, str]:
    """Return (app_name, doc_title) for two-line display."""
    # VSCode: "file.py - project - Visual Studio Code"
    vscode_match = _re.match(r"^.+ - (.+?) ?(\[.*?\])? - Visual Studio Code.*$", raw_title)
    if vscode_match:
        return "Visual Studio Code", vscode_match.group(1).strip()

    # Generic: "doc - App Name"
    parts = raw_title.rsplit(" - ", 1)
    if len(parts) == 2:
        return parts[1].strip(), parts[0].strip()

    # Fallback
    return proc if proc else raw_title, ""


class WindowItemWidget(QWidget):
    """Two-line list item: bold app name + dimmer doc title + close button."""

    close_clicked = None  # set externally after construction

    def __init__(self, app: str, doc: str, parent=None):
        super().__init__(parent)
        from PyQt6.QtWidgets import QVBoxLayout as _VBox
        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 8, 8, 8)
        layout.setSpacing(0)

        text_col = _VBox()
        text_col.setSpacing(2)

        is_vscode = app == "Visual Studio Code"
        app_color = "#4fc3f7" if is_vscode else "#f0f0f0"
        app_size = 13 if is_vscode else 11

        app_label = QLabel(app)
        app_label.setFont(QFont("Segoe UI", app_size, QFont.Weight.DemiBold))
        app_label.setStyleSheet(f"color: {app_color}; background: transparent;")
        app_label.setWordWrap(False)
        text_col.addWidget(app_label)

        if doc:
            doc_color = "#80cbc4" if is_vscode else "rgba(180, 180, 190, 200)"
            doc_size = 10 if is_vscode else 9
            doc_label = QLabel(doc)
            doc_label.setFont(QFont("Segoe UI", doc_size))
            doc_label.setStyleSheet(f"color: {doc_color}; background: transparent;")
            doc_label.setWordWrap(False)
            text_col.addWidget(doc_label)

        layout.addLayout(text_col, 1)

        self._close_btn = QLabel("✕")
        self._close_btn.setFont(QFont("Segoe UI", 10))
        self._close_btn.setStyleSheet(
            "color: rgba(180,180,190,120); background: transparent; padding: 0 6px;"
        )
        self._close_btn.setFixedWidth(28)
        self._close_btn.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._close_btn.mousePressEvent = self._on_close_press
        layout.addWidget(self._close_btn)

    def _on_close_press(self, event) -> None:
        if self.close_clicked:
            self.close_clicked()


class SwitcherWindow(QWidget):
    """The main Alt-Tab replacement overlay."""

    def __init__(self):
        super().__init__()
        self._windows: list[tuple[int, str, str]] = []
        self._poll_timer = QTimer(self)
        self._poll_timer.setInterval(30)
        self._poll_timer.timeout.connect(self._poll_alt_release)

        self.setObjectName("SwitcherBackground")
        self.setWindowTitle("LazyWinTab")
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating, True)

        # Hide from taskbar by applying WS_EX_TOOLWINDOW after the native handle exists
        GWL_EXSTYLE_SET = -20
        QTimer.singleShot(0, lambda: user32.SetWindowLongW(
            int(self.winId()), GWL_EXSTYLE_SET,
            user32.GetWindowLongW(int(self.winId()), GWL_EXSTYLE_SET) | WS_EX_TOOLWINDOW
        ))

        # Layout
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(0, 0, 0, 0)

        # Inner container (for rounded background)
        self._container = QWidget()
        self._container.setObjectName("SwitcherBackground")
        root_layout.addWidget(self._container)

        inner = QVBoxLayout(self._container)
        inner.setContentsMargins(6, 6, 6, 6)
        inner.setSpacing(0)

        header = QLabel("SWITCH TO")
        header.setObjectName("Header")
        inner.addWidget(header)

        self._list = QListWidget()
        self._list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._list.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._list.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._list.itemClicked.connect(self._on_item_clicked)
        self._list.installEventFilter(self)
        inner.addWidget(self._list)

        self.setStyleSheet(STYLESHEET)
        self.hide()

        self._hook = None
        self._hook_proc = None  # keep reference to prevent GC
        self._mouse_hook = None
        self._mouse_hook_proc = None
        self._install_hook()

    # -- Low-level keyboard hook ---------------------------------------------

    def _install_hook(self) -> None:
        """Install a WH_KEYBOARD_LL hook to intercept Alt+Tab."""
        HOOKPROC = WINFUNCTYPE(ctypes.c_long, c_int, wintypes.WPARAM, wintypes.LPARAM)

        VK_UP = 0x26
        VK_DOWN = 0x28
        VK_RETURN = 0x0D
        VK_ESCAPE = 0x1B
        VK_DELETE = 0x2E

        WM_KEYUP = 0x0101
        WM_SYSKEYUP = 0x0105

        def _hook_proc(nCode, wParam, lParam):
            if nCode >= 0:
                kb = ctypes.cast(lParam, ctypes.POINTER(wintypes.KBDLLHOOKSTRUCT)).contents
                vk = kb.vkCode

                # Alt key released while switcher is open → commit selection
                if wParam in (WM_KEYUP, WM_SYSKEYUP) and vk in (VK_MENU, VK_LMENU, VK_RMENU):
                    if self.isVisible():
                        QTimer.singleShot(0, lambda: self._dismiss(switch=True))
                    return user32.CallNextHookEx(self._hook, nCode, wParam, lParam)

            if nCode >= 0 and wParam in (WM_KEYDOWN, WM_SYSKEYDOWN):
                kb = ctypes.cast(lParam, ctypes.POINTER(wintypes.KBDLLHOOKSTRUCT)).contents
                vk = kb.vkCode
                alt_down = bool(user32.GetAsyncKeyState(VK_MENU) & 0x8000)

                if vk == VK_TAB and alt_down:
                    shift_down = bool(user32.GetAsyncKeyState(0x10) & 0x8000)
                    if shift_down:
                        QTimer.singleShot(0, self._on_shift_tab)
                    else:
                        QTimer.singleShot(0, self._on_hotkey_triggered)
                    return 1

                if self.isVisible():
                    if vk == VK_UP:
                        QTimer.singleShot(0, self._select_prev)
                        return 1
                    if vk == VK_DOWN:
                        QTimer.singleShot(0, self._select_next)
                        return 1
                    if vk == VK_RETURN:
                        QTimer.singleShot(0, lambda: self._dismiss(switch=True))
                        return 1
                    if vk == VK_ESCAPE:
                        QTimer.singleShot(0, lambda: self._dismiss(switch=False))
                        return 1
                    if vk == VK_DELETE:
                        QTimer.singleShot(0, self._close_selected)
                        return 1

            return user32.CallNextHookEx(self._hook, nCode, wParam, lParam)

        self._hook_proc = HOOKPROC(_hook_proc)
        self._hook = user32.SetWindowsHookExW(WH_KEYBOARD_LL, self._hook_proc, None, 0)
        if self._hook:
            log.info("Low-level keyboard hook installed.")
        else:
            log.error("Failed to install keyboard hook (error=%s).", ctypes.GetLastError())

        def _mouse_proc(nCode, wParam, lParam):
            if nCode >= 0 and wParam == WM_MOUSEWHEEL and self.isVisible():
                ms = ctypes.cast(lParam, ctypes.POINTER(MSLLHOOKSTRUCT)).contents
                # High word of mouseData is the wheel delta (signed)
                delta = ctypes.c_short((ms.mouseData >> 16) & 0xFFFF).value
                if delta < 0:
                    QTimer.singleShot(0, self._select_next)
                elif delta > 0:
                    QTimer.singleShot(0, self._select_prev)
            return user32.CallNextHookEx(self._mouse_hook, nCode, wParam, lParam)

        self._mouse_hook_proc = HOOKPROC(_mouse_proc)
        self._mouse_hook = user32.SetWindowsHookExW(WH_MOUSE_LL, self._mouse_hook_proc, None, 0)
        if self._mouse_hook:
            log.info("Low-level mouse hook installed.")
        else:
            log.error("Failed to install mouse hook (error=%s).", ctypes.GetLastError())

    def _remove_hook(self) -> None:
        if self._hook:
            user32.UnhookWindowsHookEx(self._hook)
            self._hook = None
            log.info("Keyboard hook removed.")
        if self._mouse_hook:
            user32.UnhookWindowsHookEx(self._mouse_hook)
            self._mouse_hook = None
            log.info("Mouse hook removed.")

    # -- Show / populate / hide ----------------------------------------------

    def _on_hotkey_triggered(self) -> None:
        if self.isVisible():
            self._select_next()
        else:
            self._populate_and_show()

    def _on_shift_tab(self) -> None:
        if self.isVisible():
            self._select_prev()
        else:
            self._populate_and_show()

    def _populate_and_show(self) -> None:
        """Enumerate windows, populate the list, and show the switcher."""
        my_hwnd = int(self.winId())
        raw = [(h, t, p) for h, t, p in enumerate_windows() if h != my_hwnd]
        self._windows = group_windows(raw)

        if not self._windows:
            log.info("No windows to switch to.")
            return

        self._list.clear()
        for idx, (hwnd, title, proc) in enumerate(self._windows):
            app, doc = _format_title(title, proc)
            item = QListWidgetItem()
            item.setSizeHint(QSize(WINDOW_WIDTH - 20, ITEM_HEIGHT))
            self._list.addItem(item)
            widget = WindowItemWidget(app, doc)
            widget.close_clicked = lambda i=idx: self._close_window_at(i)
            self._list.setItemWidget(item, widget)

        # Select the second item (first is usually the current window)
        start_index = 1 if len(self._windows) > 1 else 0
        self._list.setCurrentRow(start_index)

        # Size and center
        visible = min(len(self._windows), MAX_VISIBLE_ITEMS)
        list_height = visible * (ITEM_HEIGHT + 4) + 8
        total_height = list_height + 32  # header
        self.setFixedSize(WINDOW_WIDTH, total_height)

        screen = QApplication.primaryScreen()
        if screen:
            geo = screen.geometry()
            x = (geo.width() - WINDOW_WIDTH) // 2
            y = (geo.height() - total_height) // 2
            self.move(x, y)

        self.show()

        # Delay poll start to avoid false Alt-release detection right after hook fires
        QTimer.singleShot(200, self._poll_timer.start)
        log.info("Switcher shown with %d windows.", len(self._windows))

    def _dismiss(self, switch: bool = True) -> None:
        """Hide the switcher and optionally switch to the selected window."""
        self._poll_timer.stop()
        self.hide()

        if switch:
            row = self._list.currentRow()
            if 0 <= row < len(self._windows):
                hwnd = self._windows[row][0]
                switch_to_window(hwnd)

    # -- Navigation ----------------------------------------------------------

    def _select_next(self) -> None:
        row = self._list.currentRow()
        new_row = (row + 1) % self._list.count()
        self._list.setCurrentRow(new_row)

    def _select_prev(self) -> None:
        row = self._list.currentRow()
        new_row = (row - 1) % self._list.count()
        self._list.setCurrentRow(new_row)

    def _close_window_at(self, idx: int) -> None:
        """Close the window at the given list index and remove it from the switcher."""
        if not (0 <= idx < len(self._windows)):
            return
        hwnd = self._windows[idx][0]
        user32.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE
        self._windows.pop(idx)
        self._list.takeItem(idx)
        if not self._windows:
            self._dismiss(switch=False)
            return
        # Update close_clicked callbacks (indices shifted after removal)
        for i in range(self._list.count()):
            w = self._list.itemWidget(self._list.item(i))
            if w:
                w.close_clicked = lambda j=i: self._close_window_at(j)
        new_row = min(idx, self._list.count() - 1)
        self._list.setCurrentRow(new_row)

    def _close_selected(self) -> None:
        self._close_window_at(self._list.currentRow())

    def _on_item_clicked(self, item: QListWidgetItem) -> None:
        self._dismiss(switch=True)

    # -- Scroll wheel navigation ---------------------------------------------

    def wheelEvent(self, event) -> None:
        delta = event.angleDelta().y()
        if delta < 0:
            self._select_next()
        elif delta > 0:
            self._select_prev()
        event.accept()

    def eventFilter(self, obj, event) -> bool:
        from PyQt6.QtCore import QEvent
        if obj is self._list and event.type() == QEvent.Type.Wheel:
            delta = event.angleDelta().y()
            if delta < 0:
                self._select_next()
            elif delta > 0:
                self._select_prev()
            return True  # consume — don't scroll the list view
        return super().eventFilter(obj, event)

    # -- Poll Alt release (triggers switch) ----------------------------------

    def _poll_alt_release(self) -> None:
        """Check if the Alt key has been released."""
        alt_state = user32.GetAsyncKeyState(VK_MENU)
        if not (alt_state & 0x8000):
            # Alt released
            self._dismiss(switch=True)

    # -- Keyboard events -----------------------------------------------------

    def keyPressEvent(self, event: QKeyEvent) -> None:
        key = event.key()
        if key == Qt.Key.Key_Escape:
            self._dismiss(switch=False)
        elif key == Qt.Key.Key_Return or key == Qt.Key.Key_Enter:
            self._dismiss(switch=True)
        elif key == Qt.Key.Key_Tab:
            if event.modifiers() & Qt.KeyboardModifier.ShiftModifier:
                self._select_prev()
            else:
                self._select_next()
        elif key == Qt.Key.Key_Up:
            self._select_prev()
        elif key == Qt.Key.Key_Down:
            self._select_next()
        else:
            super().keyPressEvent(event)

    # -- Cleanup -------------------------------------------------------------

    def closeEvent(self, event) -> None:
        self._remove_hook()
        super().closeEvent(event)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

_AUTOSTART_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
_AUTOSTART_NAME = "LazyWinTab"


def _get_exe_path() -> str:
    """Return the path to use for the autostart entry (exe or python script)."""
    if getattr(sys, "frozen", False):
        return sys.executable  # PyInstaller exe
    return f'"{sys.executable}" "{os.path.abspath(__file__)}"'


def _is_autostart_enabled() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _AUTOSTART_KEY) as key:
            winreg.QueryValueEx(key, _AUTOSTART_NAME)
            return True
    except FileNotFoundError:
        return False


def _set_autostart(enable: bool) -> None:
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _AUTOSTART_KEY, access=winreg.KEY_SET_VALUE) as key:
        if enable:
            winreg.SetValueEx(key, _AUTOSTART_NAME, 0, winreg.REG_SZ, _get_exe_path())
            log.info("Autostart enabled.")
        else:
            try:
                winreg.DeleteValue(key, _AUTOSTART_NAME)
                log.info("Autostart disabled.")
            except FileNotFoundError:
                pass


INSTALL_STYLESHEET = """
QWidget#InstallerBg {
    background-color: #16161c;
    border: 1px solid rgba(255,255,255,18);
    border-radius: 16px;
}
QLabel#Title {
    color: #f0f0f0;
    font-size: 15px;
    font-weight: 600;
}
QLabel#Sub {
    color: rgba(160,160,170,200);
    font-size: 10px;
}
QLabel#Done {
    color: #69f0ae;
    font-size: 11px;
    font-weight: 600;
}
QProgressBar {
    background: rgba(255,255,255,15);
    border: none;
    border-radius: 4px;
    height: 8px;
}
QProgressBar::chunk {
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #7850dc,stop:1 #4fc3f7);
    border-radius: 4px;
}
"""


class InstallerWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setObjectName("InstallerBg")
        self.setFixedSize(360, 160)

        from PyQt6.QtWidgets import QProgressBar

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        bg = QWidget()
        bg.setObjectName("InstallerBg")
        outer.addWidget(bg)

        lay = QVBoxLayout(bg)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(10)

        title = QLabel("Installing LazyWinTab")
        title.setObjectName("Title")
        title.setFont(QFont("Segoe UI", 13, QFont.Weight.DemiBold))
        lay.addWidget(title)

        self._sub = QLabel("Setting up autostart…")
        self._sub.setObjectName("Sub")
        self._sub.setFont(QFont("Segoe UI", 9))
        lay.addWidget(self._sub)

        self._bar = QProgressBar()
        self._bar.setRange(0, 100)
        self._bar.setValue(0)
        self._bar.setTextVisible(False)
        self._bar.setFixedHeight(8)
        lay.addWidget(self._bar)

        self._done_label = QLabel("")
        self._done_label.setObjectName("Done")
        self._done_label.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        self._done_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(self._done_label)

        self.setStyleSheet(INSTALL_STYLESHEET)

        # Center on screen
        screen = QApplication.primaryScreen()
        if screen:
            geo = screen.geometry()
            self.move((geo.width() - self.width()) // 2, (geo.height() - self.height()) // 2)

        self._step = 0
        self._timer = QTimer(self)
        self._timer.setInterval(30)
        self._timer.timeout.connect(self._tick)
        QTimer.singleShot(200, self._timer.start)

    def _tick(self):
        self._step += 2
        self._bar.setValue(min(self._step, 100))

        if self._step == 40:
            self._sub.setText("Writing registry entry…")
        elif self._step == 70:
            _set_autostart(True)
            self._sub.setText("Finalizing…")
        elif self._step >= 100:
            self._timer.stop()
            self._bar.setValue(100)
            self._sub.setText("")
            self._done_label.setText("LazyWinTab installed! Starts with Windows.")
            QTimer.singleShot(2200, self.close)


def main():
    is_install = "--install" in sys.argv or not _is_autostart_enabled()

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    if is_install:
        installer = InstallerWindow()
        installer.show()
        # After installer closes, continue to run normally
        installer.destroyed.connect(lambda: None)  # keep ref

    try:
        switcher = SwitcherWindow()
    except Exception:
        import traceback
        traceback.print_exc()
        input("Press Enter to exit")
        return

    # System tray icon keeps the event loop alive and gives a way to quit.
    tray = QSystemTrayIcon(app)
    tray.setIcon(app.style().standardIcon(app.style().StandardPixmap.SP_ComputerIcon))
    tray_menu = QMenu()

    autostart_action = tray_menu.addAction("Run on startup")
    autostart_action.setCheckable(True)
    autostart_action.setChecked(_is_autostart_enabled())

    def _toggle_autostart(checked: bool):
        _set_autostart(checked)

    autostart_action.toggled.connect(_toggle_autostart)

    tray_menu.addSeparator()
    quit_action = tray_menu.addAction("Quit LazyWinTab")
    quit_action.triggered.connect(app.quit)
    tray.setContextMenu(tray_menu)
    tray.setToolTip("LazyWinTab — Alt+Tab switcher running")
    tray.show()

    log.info("LazyWinTab is running. Press Alt+Tab to use the custom switcher.")
    log.info("Right-click the tray icon to quit.")

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
