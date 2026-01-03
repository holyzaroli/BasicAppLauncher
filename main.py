import sys
import os
import json
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QListWidget, QLineEdit, QAbstractItemView
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QFontDatabase, QColor, QPalette
from PySide6.QtWidgets import QStyleFactory

import winreg
import win32gui
import win32con

# For UWP app enumeration
import win32com.client

# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------

def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_workerw():
    progman = win32gui.FindWindow("Progman", None)
    win32gui.SendMessageTimeout(
        progman,
        0x052C,
        0,
        0,
        win32con.SMTO_NORMAL,
        100
    )

    workerw = None

    def enum_windows(hwnd, _):
        nonlocal workerw
        if win32gui.GetClassName(hwnd) == "WorkerW":
            shell = win32gui.FindWindowEx(hwnd, 0, "SHELLDLL_DefView", None)
            if shell:
                workerw = hwnd

    win32gui.EnumWindows(enum_windows, None)
    return workerw

# ------------------------------------------------------------
# Config
# ------------------------------------------------------------

class ConfigManager:
    def __init__(self):
        self.config_path = os.path.join(app_dir(), "config.json")
        self.default_config = {
            "background_color": "#1E1E1E",
            "text_color": "#FFFFFF",
            "font_path": "C:/Windows/Fonts/Consolas.ttf",
            "font_size": 12,
            "auto_startup": True
        }
        self.config = self.load()

    def load(self):
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        self.save(self.default_config)
        return self.default_config

    def save(self, cfg):
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4)

    def get(self, key, default=None):
        return self.config.get(key, default)

# ------------------------------------------------------------
# Main Window
# ------------------------------------------------------------

class ApplicationLauncher(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = ConfigManager()

        self.start_menu_dirs = [
            Path(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"),
            Path(os.environ["APPDATA"]) / "Microsoft/Windows/Start Menu/Programs"
        ]

        self.items = []  # list of tuples (name, path)
        self.setup_ui()
        self.apply_font()
        self.load_items()
        self.setup_startup()

    # ----------------- UI -----------------
    def setup_ui(self):
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool)
        central = QWidget()
        self.setCentralWidget(central)

        layout = QVBoxLayout(central)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        self.search = QLineEdit()
        self.search.setPlaceholderText("Search...")
        self.search.textChanged.connect(self.filter_items)
        layout.addWidget(self.search)

        self.list = QListWidget()
        self.list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.list.itemActivated.connect(self.launch_selected)
        self.list.setUniformItemSizes(True)
        layout.addWidget(self.list)

        bg = QColor(self.config.get("background_color"))
        fg = QColor(self.config.get("text_color"))

        palette = QPalette()
        palette.setColor(QPalette.Window, bg)
        palette.setColor(QPalette.Base, bg)
        palette.setColor(QPalette.Text, fg)
        palette.setColor(QPalette.WindowText, fg)
        self.setPalette(palette)
        self.setAutoFillBackground(True)

    # ----------------- Font -----------------
    def apply_font(self):
        font_path = self.config.get("font_path")
        font_size = self.config.get("font_size", 12)
        families = []

        if font_path:
            font_path = (
                os.path.join(app_dir(), font_path)
                if not os.path.isabs(font_path)
                else font_path
            )
            if os.path.isfile(font_path):
                fid = QFontDatabase.addApplicationFont(font_path)
                if fid != -1:
                    families = QFontDatabase.applicationFontFamilies(fid)

        if families:
            font = QFont(families[0], font_size)
        else:
            font = QFont("Consolas", font_size)
            font.setStyleHint(QFont.Monospace)

        self.setFont(font)
        self.search.setFont(font)
        self.list.setFont(font)

    # ----------------- Load apps -----------------
    def load_items(self):
        self.items.clear()
        self.list.clear()
        seen = set()

        # --- Classic Start Menu .lnk apps ---
        for base in self.start_menu_dirs:
            if not base.exists():
                continue
            for item in base.rglob("*.lnk"):
                name = item.stem
                key = name.lower()
                if key in seen:
                    continue
                seen.add(key)
                self.items.append((name, str(item)))

        # --- UWP / Microsoft Store apps ---
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            apps_folder = shell.NameSpace("shell:AppsFolder")
            for item in apps_folder.Items():
                name = item.Name
                key = name.lower()
                if key in seen:
                    continue
                seen.add(key)
                path = f"shell:appsFolder\\{item.Path}"
                self.items.append((name, path))
        except Exception as e:
            print("Warning: Could not load UWP apps:", e)

        self.items.sort(key=lambda x: x[0].lower())
        for name, _ in self.items:
            self.list.addItem(name)

        if self.list.count():
            self.list.setCurrentRow(0)

    # ----------------- Search -----------------
    def filter_items(self, text):
        text = text.lower()
        self.list.clear()
        for name, _ in self.items:
            if text in name.lower():
                self.list.addItem(name)

    # ----------------- Launch -----------------
    def launch_selected(self):
        item = self.list.currentItem()
        if not item:
            return
        name = item.text()
        for n, path in self.items:
            if n == name:
                try:
                    os.startfile(path)
                except Exception as e:
                    print(f"Failed to launch {name}: {e}")
                break

    # ----------------- Startup -----------------
    def setup_startup(self):
        if not self.config.get("auto_startup", False):
            return
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_SET_VALUE
            )
            winreg.SetValueEx(
                key,
                "ConsoleLauncher",
                0,
                winreg.REG_SZ,
                sys.executable
            )
            winreg.CloseKey(key)
        except Exception:
            pass

    # ----------------- Attach to desktop -----------------
    def attach_to_desktop(self):
        hwnd = int(self.winId())
        workerw = get_workerw()
        if workerw:
            win32gui.SetParent(hwnd, workerw)

# ------------------------------------------------------------
# Entry
# ------------------------------------------------------------

def main():
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create("Fusion"))

    launcher = ApplicationLauncher()
    screen = QApplication.primaryScreen().availableGeometry()
    launcher.setGeometry(screen)
    launcher.show()
    launcher.attach_to_desktop()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
