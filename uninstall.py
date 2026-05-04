"""
LazyWinTab Uninstaller — removes autostart registry entry and config data.
"""

import sys
import os
import shutil
import pathlib
import winreg

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QProgressBar,
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_AUTOSTART_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
_AUTOSTART_NAME = "LazyWinTab"

STYLESHEET = """
QWidget#Root {
    background-color: #16161c;
    border: 1px solid rgba(255,255,255,18);
    border-radius: 16px;
}
QLabel#Title {
    color: #f0f0f0;
    font-size: 14px;
    font-weight: 600;
}
QLabel#Body {
    color: rgba(180,180,190,210);
    font-size: 10px;
}
QLabel#Done {
    color: #69f0ae;
    font-size: 11px;
    font-weight: 600;
}
QLabel#Err {
    color: #ef5350;
    font-size: 10px;
}
QPushButton#UninstallBtn {
    background: #c62828;
    border: none;
    border-radius: 8px;
    color: #fff;
    font-size: 11px;
    font-weight: 700;
    padding: 7px 22px;
}
QPushButton#UninstallBtn:hover {
    background: #e53935;
}
QPushButton#CancelBtn {
    background: transparent;
    border: 1px solid rgba(255,255,255,20);
    border-radius: 8px;
    color: rgba(180,180,190,200);
    font-size: 11px;
    padding: 7px 16px;
}
QPushButton#CancelBtn:hover {
    border: 1px solid rgba(255,255,255,50);
    color: #e0e0e0;
}
QProgressBar {
    background: rgba(255,255,255,15);
    border: none;
    border-radius: 4px;
}
QProgressBar::chunk {
    background: #c62828;
    border-radius: 4px;
}
"""

# ---------------------------------------------------------------------------
# Uninstall logic
# ---------------------------------------------------------------------------

def _remove_autostart() -> None:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _AUTOSTART_KEY,
                            access=winreg.KEY_SET_VALUE) as key:
            winreg.DeleteValue(key, _AUTOSTART_NAME)
    except FileNotFoundError:
        pass


def _remove_config() -> None:
    config_dir = pathlib.Path(os.environ["APPDATA"]) / "LazyWinTab"
    if config_dir.exists():
        shutil.rmtree(config_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

class UninstallWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Uninstall LazyWinTab")
        self.setObjectName("Root")
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(380, 210)
        self.setStyleSheet(STYLESHEET)

        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(12)

        title = QLabel("Uninstall LazyWinTab")
        title.setObjectName("Title")
        title.setFont(QFont("Segoe UI", 13, QFont.Weight.DemiBold))
        root.addWidget(title)

        body = QLabel(
            "This will remove the autostart registry entry\n"
            "and delete all settings data."
        )
        body.setObjectName("Body")
        body.setFont(QFont("Segoe UI", 9))
        root.addWidget(body)

        self._bar = QProgressBar()
        self._bar.setRange(0, 100)
        self._bar.setValue(0)
        self._bar.setTextVisible(False)
        self._bar.setFixedHeight(6)
        self._bar.hide()
        root.addWidget(self._bar)

        self._status = QLabel("")
        self._status.setObjectName("Done")
        self._status.setFont(QFont("Segoe UI", 10, QFont.Weight.DemiBold))
        self._status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._status.hide()
        root.addWidget(self._status)

        root.addStretch()

        btn_row = QHBoxLayout()
        self._cancel_btn = QPushButton("Cancel")
        self._cancel_btn.setObjectName("CancelBtn")
        self._cancel_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._cancel_btn.clicked.connect(self.close)
        btn_row.addWidget(self._cancel_btn)
        btn_row.addStretch()
        self._uninstall_btn = QPushButton("Uninstall")
        self._uninstall_btn.setObjectName("UninstallBtn")
        self._uninstall_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._uninstall_btn.clicked.connect(self._start_uninstall)
        btn_row.addWidget(self._uninstall_btn)
        root.addLayout(btn_row)

        # Center on screen
        screen = QApplication.primaryScreen()
        if screen:
            geo = screen.geometry()
            self.move(
                (geo.width() - self.width()) // 2,
                (geo.height() - self.height()) // 2,
            )

        self._step = 0

    def _start_uninstall(self) -> None:
        self._uninstall_btn.setEnabled(False)
        self._cancel_btn.setEnabled(False)
        self._bar.show()
        self._status.hide()
        self._step = 0
        self._timer = QTimer(self)
        self._timer.setInterval(25)
        self._timer.timeout.connect(self._tick)
        self._timer.start()

    def _tick(self) -> None:
        self._step += 3
        self._bar.setValue(min(self._step, 100))

        if self._step == 30:
            _remove_autostart()
        elif self._step == 65:
            _remove_config()
        elif self._step >= 100:
            self._timer.stop()
            self._bar.setValue(100)
            self._status.setObjectName("Done")
            self._status.setText("Uninstalled successfully.")
            self._status.setStyleSheet("color: #69f0ae; font-size: 11px; font-weight: 600;")
            self._status.show()
            QTimer.singleShot(2000, self.close)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)
    win = UninstallWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
