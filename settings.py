"""
LazyWinTab Settings — per-app color editor.
Run standalone or launched from the tray icon.
"""

import sys
import os
import json
import pathlib
import subprocess

from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QColor, QFont, QIcon
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QScrollArea,
    QFrame, QColorDialog, QMessageBox,
)

# ---------------------------------------------------------------------------
# Config helpers (mirrors main.py)
# ---------------------------------------------------------------------------

_DEFAULT_COLORS: dict[str, str] = {
    "Visual Studio Code": "#4fc3f7",
}


def _config_path() -> pathlib.Path:
    d = pathlib.Path(os.environ["APPDATA"]) / "LazyWinTab"
    d.mkdir(exist_ok=True)
    return d / "colors.json"


def _load_colors() -> dict[str, str]:
    try:
        return json.loads(_config_path().read_text(encoding="utf-8"))
    except Exception:
        return dict(_DEFAULT_COLORS)


def _save_colors(colors: dict[str, str]) -> None:
    _config_path().write_text(
        json.dumps(colors, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def _uninstall_exe() -> list[str]:
    if getattr(sys, "frozen", False):
        exe = pathlib.Path(sys.executable).parent / "LazyWinTabUninstall.exe"
        return [str(exe)]
    return [sys.executable, str(pathlib.Path(__file__).parent / "uninstall.py")]


# ---------------------------------------------------------------------------
# Stylesheet
# ---------------------------------------------------------------------------

STYLESHEET = """
QWidget#Root {
    background-color: #16161c;
}
QWidget#Card {
    background-color: rgba(255,255,255,6);
    border-radius: 10px;
}
QLabel#Title {
    color: #f0f0f0;
    font-size: 15px;
    font-weight: 600;
}
QLabel#Section {
    color: rgba(160,160,170,180);
    font-size: 9px;
    letter-spacing: 1px;
}
QLabel#AppName {
    color: #e8e8e8;
    font-size: 11px;
}
QLabel#Hint {
    color: rgba(140,140,150,180);
    font-size: 9px;
}
QPushButton#Swatch {
    border-radius: 6px;
    border: 1px solid rgba(255,255,255,30);
    min-width: 52px;
    max-width: 52px;
    min-height: 22px;
    max-height: 22px;
}
QPushButton#Swatch:hover {
    border: 1px solid rgba(255,255,255,80);
}
QPushButton#Remove {
    color: rgba(180,80,80,200);
    background: transparent;
    border: none;
    font-size: 14px;
    padding: 0;
    min-width: 20px;
    max-width: 20px;
}
QPushButton#Remove:hover {
    color: #ef5350;
}
QLineEdit#AddInput {
    background: rgba(255,255,255,8);
    border: 1px solid rgba(255,255,255,20);
    border-radius: 6px;
    color: #e0e0e0;
    font-size: 11px;
    padding: 4px 8px;
    selection-background-color: rgba(120,80,220,160);
}
QLineEdit#AddInput:focus {
    border: 1px solid rgba(120,80,220,200);
}
QPushButton#AddBtn {
    background: rgba(120,80,220,180);
    border: none;
    border-radius: 6px;
    color: #fff;
    font-size: 11px;
    font-weight: 600;
    padding: 4px 14px;
}
QPushButton#AddBtn:hover {
    background: rgba(140,100,240,200);
}
QPushButton#SaveBtn {
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #7850dc,stop:1 #4fc3f7);
    border: none;
    border-radius: 8px;
    color: #fff;
    font-size: 11px;
    font-weight: 700;
    padding: 7px 28px;
}
QPushButton#SaveBtn:hover {
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #9070f0,stop:1 #60d0ff);
}
QPushButton#ResetBtn {
    background: transparent;
    border: 1px solid rgba(255,255,255,20);
    border-radius: 8px;
    color: rgba(180,180,190,200);
    font-size: 11px;
    padding: 7px 16px;
}
QPushButton#ResetBtn:hover {
    border: 1px solid rgba(255,255,255,50);
    color: #e0e0e0;
}
QPushButton#UninstallBtn {
    background: transparent;
    border: 1px solid rgba(180,50,50,120);
    border-radius: 8px;
    color: rgba(220,80,80,200);
    font-size: 11px;
    padding: 7px 16px;
}
QPushButton#UninstallBtn:hover {
    border: 1px solid #ef5350;
    color: #ef5350;
}
QScrollArea {
    background: transparent;
    border: none;
}
QScrollBar:vertical {
    background: transparent;
    width: 4px;
}
QScrollBar::handle:vertical {
    background: rgba(255,255,255,40);
    border-radius: 2px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
"""

# ---------------------------------------------------------------------------
# Color row widget
# ---------------------------------------------------------------------------

class ColorRow(QWidget):
    def __init__(self, name: str, color: str, on_remove, parent=None):
        super().__init__(parent)
        self.name = name
        self._color = color

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 6, 12, 6)
        layout.setSpacing(10)

        lbl = QLabel(name)
        lbl.setObjectName("AppName")
        lbl.setFont(QFont("Segoe UI", 11))
        layout.addWidget(lbl, 1)

        self._swatch = QPushButton()
        self._swatch.setObjectName("Swatch")
        self._swatch.setCursor(Qt.CursorShape.PointingHandCursor)
        self._swatch.setToolTip("Click to change color")
        self._apply_color(color)
        self._swatch.clicked.connect(self._pick_color)
        layout.addWidget(self._swatch)

        remove_btn = QPushButton("✕")
        remove_btn.setObjectName("Remove")
        remove_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        remove_btn.setToolTip("Remove")
        remove_btn.clicked.connect(on_remove)
        layout.addWidget(remove_btn)

    def _apply_color(self, color: str) -> None:
        self._color = color
        self._swatch.setStyleSheet(
            f"QPushButton#Swatch {{ background-color: {color}; "
            f"border-radius: 6px; border: 1px solid rgba(255,255,255,30); }}"
            f"QPushButton#Swatch:hover {{ border: 1px solid rgba(255,255,255,80); }}"
        )

    def _pick_color(self) -> None:
        initial = QColor(self._color)
        picked = QColorDialog.getColor(initial, self, f"Color for {self.name}")
        if picked.isValid():
            self._apply_color(picked.name())

    def current_color(self) -> str:
        return self._color


# ---------------------------------------------------------------------------
# Main settings window
# ---------------------------------------------------------------------------

class SettingsWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("LazyWinTab Settings")
        self.setMinimumWidth(440)
        self.setObjectName("Root")
        self.setStyleSheet(STYLESHEET)

        self._rows: list[ColorRow] = []

        root = QVBoxLayout(self)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(16)

        # Title
        title = QLabel("LazyWinTab Settings")
        title.setObjectName("Title")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.DemiBold))
        root.addWidget(title)

        # Section label
        section = QLabel("APP COLORS")
        section.setObjectName("Section")
        section.setFont(QFont("Segoe UI", 8))
        root.addWidget(section)

        # Scrollable list
        self._list_widget = QWidget()
        self._list_layout = QVBoxLayout(self._list_widget)
        self._list_layout.setContentsMargins(0, 0, 0, 0)
        self._list_layout.setSpacing(2)
        self._list_layout.addStretch()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self._list_widget)
        scroll.setMinimumHeight(200)
        scroll.setMaximumHeight(340)
        root.addWidget(scroll)

        # Add row
        add_row = QHBoxLayout()
        add_row.setSpacing(8)
        self._add_input = QLineEdit()
        self._add_input.setObjectName("AddInput")
        self._add_input.setPlaceholderText("App name (e.g. chrome, Notepad++)")
        self._add_input.setFont(QFont("Segoe UI", 10))
        self._add_input.returnPressed.connect(self._add_entry)
        add_row.addWidget(self._add_input, 1)
        add_btn = QPushButton("Add")
        add_btn.setObjectName("AddBtn")
        add_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        add_btn.clicked.connect(self._add_entry)
        add_row.addWidget(add_btn)
        root.addLayout(add_row)

        hint = QLabel("Match by app display name or process name (without .exe)")
        hint.setObjectName("Hint")
        hint.setFont(QFont("Segoe UI", 8))
        root.addWidget(hint)

        # Separator
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: rgba(255,255,255,15);")
        root.addWidget(sep)

        # Bottom buttons
        btn_row = QHBoxLayout()
        reset_btn = QPushButton("Reset to defaults")
        reset_btn.setObjectName("ResetBtn")
        reset_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        reset_btn.clicked.connect(self._reset)
        btn_row.addWidget(reset_btn)

        uninstall_btn = QPushButton("Uninstall…")
        uninstall_btn.setObjectName("UninstallBtn")
        uninstall_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        uninstall_btn.clicked.connect(lambda: subprocess.Popen(_uninstall_exe()))
        btn_row.addWidget(uninstall_btn)

        btn_row.addStretch()
        save_btn = QPushButton("Save")
        save_btn.setObjectName("SaveBtn")
        save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        save_btn.clicked.connect(self._save)
        btn_row.addWidget(save_btn)
        root.addLayout(btn_row)

        self._load()
        self.adjustSize()

    # -- Data ----------------------------------------------------------------

    def _load(self) -> None:
        colors = _load_colors()
        for name, color in colors.items():
            self._insert_row(name, color)

    def _insert_row(self, name: str, color: str) -> None:
        row = ColorRow(name, color, on_remove=lambda checked=False, r=None: self._remove_row(r))
        # Fix closure: bind row reference after creation
        row.findChild(QPushButton, "Remove")
        # Re-wire with correct reference
        btns = [c for c in row.children() if isinstance(c, QPushButton) and c.objectName() == "Remove"]
        if btns:
            btns[0].clicked.disconnect()
            btns[0].clicked.connect(lambda checked=False, r=row: self._remove_row(r))

        self._rows.append(row)
        # Insert before the stretch (last item)
        idx = self._list_layout.count() - 1
        self._list_layout.insertWidget(idx, row)

    def _remove_row(self, row: ColorRow) -> None:
        self._rows.remove(row)
        self._list_layout.removeWidget(row)
        row.deleteLater()

    def _add_entry(self) -> None:
        name = self._add_input.text().strip()
        if not name:
            return
        # Don't add duplicates
        if any(r.name.lower() == name.lower() for r in self._rows):
            self._add_input.clear()
            return
        self._insert_row(name, "#f0f0f0")
        self._add_input.clear()

    def _reset(self) -> None:
        for row in list(self._rows):
            self._remove_row(row)
        for name, color in _DEFAULT_COLORS.items():
            self._insert_row(name, color)

    def _save(self) -> None:
        colors = {row.name: row.current_color() for row in self._rows}
        _save_colors(colors)
        self.close()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)
    win = SettingsWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()