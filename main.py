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
import ctypes
import ctypes.wintypes as wintypes
import logging
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
user32.SetWindowsHookExW.argtypes = [c_int, ctypes.c_void_p, wintypes.HINSTANCE, wintypes.DWORD]
user32.SetWindowsHookExW.restype = ctypes.c_void_p
user32.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]
user32.UnhookWindowsHookEx.restype = wintypes.BOOL
user32.CallNextHookEx.argtypes = [ctypes.c_void_p, c_int, wintypes.WPARAM, wintypes.LPARAM]
user32.CallNextHookEx.restype = ctypes.c_long

WH_KEYBOARD_LL = 13
WM_KEYDOWN = 0x0100
WM_SYSKEYDOWN = 0x0104


class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", wintypes.DWORD),
        ("scanCode", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


wintypes.KBDLLHOOKSTRUCT = KBDLLHOOKSTRUCT

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


def enumerate_windows() -> list[tuple[int, str]]:
    """Return a list of (hwnd, title) for all Alt-Tab eligible windows."""
    results: list[tuple[int, str]] = []

    @ENUMWINDOWSPROC
    def callback(hwnd, _lparam):
        if _is_alt_tab_window(hwnd):
            title = _get_window_text(hwnd)
            results.append((hwnd, title))
        return True

    user32.EnumWindows(callback, 0)
    return results


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
    background-color: rgba(30, 30, 30, 240);
    border: 1px solid rgba(80, 80, 80, 200);
    border-radius: 12px;
}

QListWidget {
    background: transparent;
    border: none;
    outline: none;
    padding: 4px;
}

QListWidget::item {
    background: transparent;
    border-radius: 8px;
    padding: 8px 12px;
    margin: 2px 4px;
    color: #e0e0e0;
}

QListWidget::item:selected {
    background-color: rgba(60, 120, 200, 180);
    color: #ffffff;
}

QListWidget::item:hover {
    background-color: rgba(70, 70, 70, 150);
}

QListWidget::item:selected:hover {
    background-color: rgba(60, 120, 200, 200);
}

QLabel#Title {
    color: rgba(180, 180, 180, 220);
    font-size: 11px;
    padding: 6px 12px 2px 12px;
}
"""

ITEM_HEIGHT = 48
MAX_VISIBLE_ITEMS = 12
WINDOW_WIDTH = 420


class WindowItemWidget(QWidget):
    """Custom widget rendered for each list item (icon + title)."""

    def __init__(self, title: str, icon_pixmap: QPixmap | None, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(10)

        icon_label = QLabel()
        if icon_pixmap and not icon_pixmap.isNull():
            icon_label.setPixmap(icon_pixmap.scaled(
                28, 28, Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            ))
        else:
            # Fallback: draw a simple colored square
            pm = QPixmap(28, 28)
            pm.fill(QColor(100, 100, 100))
            icon_label.setPixmap(pm)
        icon_label.setFixedSize(28, 28)
        layout.addWidget(icon_label)

        title_label = QLabel(title)
        title_label.setFont(QFont("Segoe UI", 11))
        title_label.setStyleSheet("color: #e0e0e0; background: transparent;")
        title_label.setWordWrap(False)
        layout.addWidget(title_label, 1)


class SwitcherWindow(QWidget):
    """The main Alt-Tab replacement overlay."""

    def __init__(self):
        super().__init__()
        self._windows: list[tuple[int, str]] = []
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
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating, False)

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

        header = QLabel("Switch to")
        header.setObjectName("Title")
        inner.addWidget(header)

        self._list = QListWidget()
        self._list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._list.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._list.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._list.itemClicked.connect(self._on_item_clicked)
        inner.addWidget(self._list)

        self.setStyleSheet(STYLESHEET)
        self.hide()

        self._hook = None
        self._hook_proc = None  # keep reference to prevent GC
        self._install_hook()

    # -- Low-level keyboard hook ---------------------------------------------

    def _install_hook(self) -> None:
        """Install a WH_KEYBOARD_LL hook to intercept Alt+Tab."""
        HOOKPROC = WINFUNCTYPE(ctypes.c_long, c_int, wintypes.WPARAM, wintypes.LPARAM)

        def _hook_proc(nCode, wParam, lParam):
            if nCode >= 0 and wParam in (WM_KEYDOWN, WM_SYSKEYDOWN):
                kb = ctypes.cast(lParam, ctypes.POINTER(wintypes.KBDLLHOOKSTRUCT)).contents
                if kb.vkCode == VK_TAB:
                    alt_down = bool(user32.GetAsyncKeyState(VK_MENU) & 0x8000)
                    if alt_down:
                        # Schedule on Qt thread to avoid reentrancy issues
                        QTimer.singleShot(0, self._on_hotkey_triggered)
                        return 1  # suppress the keystroke
            return user32.CallNextHookEx(self._hook, nCode, wParam, lParam)

        self._hook_proc = HOOKPROC(_hook_proc)
        self._hook = user32.SetWindowsHookExW(WH_KEYBOARD_LL, self._hook_proc, None, 0)
        if self._hook:
            log.info("Low-level keyboard hook installed.")
        else:
            log.error("Failed to install keyboard hook (error=%s).", ctypes.GetLastError())

    def _remove_hook(self) -> None:
        if self._hook:
            user32.UnhookWindowsHookEx(self._hook)
            self._hook = None
            log.info("Keyboard hook removed.")

    # -- Show / populate / hide ----------------------------------------------

    def _on_hotkey_triggered(self) -> None:
        """Called when Alt+Tab is pressed globally."""
        if self.isVisible():
            # Already open: cycle to next
            self._select_next()
        else:
            self._populate_and_show()

    def _populate_and_show(self) -> None:
        """Enumerate windows, populate the list, and show the switcher."""
        self._windows = enumerate_windows()
        # Filter out our own window
        my_hwnd = int(self.winId())
        self._windows = [(h, t) for h, t in self._windows if h != my_hwnd]

        if not self._windows:
            log.info("No windows to switch to.")
            return

        self._list.clear()
        for hwnd, title in self._windows:
            icon_pm = get_window_icon(hwnd, 32)
            item = QListWidgetItem()
            item.setSizeHint(QSize(WINDOW_WIDTH - 20, ITEM_HEIGHT))
            self._list.addItem(item)
            widget = WindowItemWidget(title, icon_pm)
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
        self.activateWindow()
        self.raise_()

        # Start polling for Alt key release
        self._poll_timer.start()
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

    def _on_item_clicked(self, item: QListWidgetItem) -> None:
        self._dismiss(switch=True)

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

def main():
    print("Starting LazyWinTab...", flush=True)
    app = QApplication(sys.argv)
    print("QApplication created", flush=True)
    app.setQuitOnLastWindowClosed(False)

    try:
        switcher = SwitcherWindow()
        print("SwitcherWindow created", flush=True)
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("Press Enter to exit")
        return

    # System tray icon keeps the event loop alive and gives a way to quit.
    tray = QSystemTrayIcon(app)
    tray.setIcon(app.style().standardIcon(app.style().StandardPixmap.SP_ComputerIcon))
    tray_menu = QMenu()
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
