import sys
import zipfile
import io
import shutil
import xml.etree.ElementTree as ET
import winreg
import os
import json
import subprocess
import ctypes  # Required for Taskbar Icon Fix
from pathlib import Path

# --- DEPENDENCY CHECK ---
try:
    import zipfile_zstd
    HAS_ZSTD = True
except ImportError:
    HAS_ZSTD = False

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QLineEdit, 
                             QScrollArea, QFrame, QGraphicsDropShadowEffect, 
                             QLayout, QToolTip, QStackedWidget, QFileDialog, QMessageBox, QStackedLayout)
from PyQt6.QtCore import (Qt, QSize, QPropertyAnimation, QEasingCurve, 
                          pyqtSignal, QRect, QPoint, pyqtProperty)
from PyQt6.QtGui import QColor, QFont, QPainter, QFontMetrics, QPixmap, QCursor, QIcon

# --- CONFIGURATION ---
CONFIG_FILE = "mod_config.json"
APP_ID = u'wartype.hon.modmanager.v0.5.8' # Arbitrary unique ID for Taskbar Grouping

THEME = {
    "bg_primary": "#090909",       
    "bg_secondary": "#141414",     
    "bg_tertiary": "#202020",      
    "accent": "#ff4d4d",            
    "accent_hover": "#ff6666",
    "success": "#4dff88",        
    "success_hover": "#66ff99",
    "text_primary": "#ffffff",     
    "text_secondary": "#9e9e9e",   
    "border": "#2c2c2c",           
    "btn_inactive": "#1f1f1f"      
}

INSTALLED_MODS = [] 

# --- UTILS ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class NullWriter:
    """A file-like object that discards all data written to it."""
    def write(self, text): pass
    def flush(self): pass
    def isatty(self): return False

# --- LAYOUTS & WIDGETS ---

class FlowLayout(QLayout):
    def __init__(self, parent=None, margin=0, h_spacing=15, v_spacing=15):
        super(FlowLayout, self).__init__(parent)
        self.h_spacing = h_spacing
        self.v_spacing = v_spacing
        self.setContentsMargins(margin, margin, margin, margin)
        self._item_list = []

    def __del__(self):
        item = self.takeAt(0)
        while item:
            item = self.takeAt(0)

    def addItem(self, item):
        self._item_list.append(item)

    def count(self):
        return len(self._item_list)

    def itemAt(self, index):
        if 0 <= index < len(self._item_list):
            return self._item_list[index]
        return None

    def takeAt(self, index):
        if 0 <= index < len(self._item_list):
            return self._item_list.pop(index)
        return None

    def expandingDirections(self):
        return Qt.Orientation(0)

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, width):
        height = self._do_layout(QRect(0, 0, width, 0), True)
        return height

    def setGeometry(self, rect):
        super(FlowLayout, self).setGeometry(rect)
        self._do_layout(rect, False)

    def sizeHint(self):
        return self.minimumSize()

    def minimumSize(self):
        size = QSize()
        for item in self._item_list:
            size = size.expandedTo(item.minimumSize())
        size += QSize(2 * self.contentsMargins().top(), 2 * self.contentsMargins().top())
        return size

    def _do_layout(self, rect, test_only):
        x = rect.x()
        y = rect.y()
        line_height = 0

        for item in self._item_list:
            wid = item.widget()
            space_x = self.h_spacing
            space_y = self.v_spacing
            next_x = x + item.sizeHint().width() + space_x
            if next_x - space_x > rect.right() and line_height > 0:
                x = rect.x()
                y = y + line_height + space_y
                next_x = x + item.sizeHint().width() + space_x
                line_height = 0
            if not test_only:
                item.setGeometry(QRect(QPoint(x, y), item.sizeHint()))
            x = next_x
            line_height = max(line_height, item.sizeHint().height())
        return y + line_height - rect.y()

class SidebarButton(QPushButton):
    def __init__(self, text, icon_char, is_locked=False):
        super().__init__()
        self.is_locked = is_locked
        self.original_text = text
        
        if not self.is_locked:
            self.setCheckable(True)
            self.setCursor(Qt.CursorShape.PointingHandCursor)
        else:
            self.setCheckable(False)
            self.setCursor(Qt.CursorShape.ForbiddenCursor)
            
        self.setFixedHeight(45)
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(20, 0, 10, 0)
        layout.setSpacing(15)
        
        self.icon_lbl = QLabel(icon_char)
        self.icon_lbl.setFixedSize(24, 24)
        self.icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.text_lbl = QLabel(text)
        self.text_lbl.setStyleSheet("font-weight: 600;")
        
        layout.addWidget(self.icon_lbl)
        layout.addWidget(self.text_lbl)
        layout.addStretch()

        self.update_style()

    def update_style(self):
        if self.is_locked:
            # Default "Locked" style (Normal state)
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {THEME['btn_inactive']};
                    border: 1px solid {THEME['border']};
                    border-radius: 6px;
                    margin: 2px 10px;
                    text-align: left;
                    color: {THEME['text_secondary']};
                }}
                QLabel {{
                    color: {THEME['text_secondary']}; 
                    background: transparent;
                    border: none;
                    font-size: 14px;
                }}
            """)
        else:
            # Standard Active Style
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {THEME['btn_inactive']};
                    border: 1px solid {THEME['border']};
                    border-radius: 6px;
                    margin: 2px 10px;
                    text-align: left;
                }}
                QPushButton:hover {{
                    background-color: {THEME['bg_tertiary']};
                    border: 1px solid {THEME['accent']};
                }}
                QPushButton:checked {{
                    background-color: rgba(255, 77, 77, 0.15);
                    border: 1px solid {THEME['accent']};
                    border-left: 4px solid {THEME['accent']};
                }}
                QLabel {{
                    color: #ffffff; 
                    background: transparent;
                    border: none;
                    font-size: 14px;
                }}
            """)

    def enterEvent(self, event):
        if self.is_locked:
            self.text_lbl.setText("Coming Soon")
            # Greyed out / Disabled look on Hover
            self.setStyleSheet(f"""
                QPushButton {{
                    background-color: {THEME['bg_primary']};
                    border: 1px dashed {THEME['text_secondary']};
                    border-radius: 6px;
                    margin: 2px 10px;
                    text-align: left;
                }}
                QLabel {{
                    color: {THEME['text_secondary']}; 
                    background: transparent;
                    border: none;
                    font-size: 14px;
                    font-style: italic;
                }}
            """)
        super().enterEvent(event)

    def leaveEvent(self, event):
        if self.is_locked:
            self.text_lbl.setText(self.original_text)
            self.update_style()
        super().leaveEvent(event)

class ModernButton(QPushButton):
    def __init__(self, text, is_primary=False):
        super().__init__(text)
        self.is_primary = is_primary
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedHeight(36)
        self.update_style()

    def set_launch_mode(self, is_launch):
        if is_launch:
            self.setText("Launch HoN")
            bg = THEME["success"]
            hover = THEME["success_hover"]
            text_color = "#000000"
            border = THEME["success"]
        else:
            self.setText("Apply Changes")
            bg = THEME["accent"]
            hover = THEME["accent_hover"]
            text_color = THEME["bg_primary"]
            border = THEME["accent"]
            
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {bg};
                color: {text_color};
                border: 1px solid {border};
                border-radius: 4px;
                padding: 0 16px;
                font-weight: bold;
                font-family: 'Segoe UI', sans-serif;
            }}
            QPushButton:hover {{
                background-color: {hover};
                border-color: {hover};
            }}
        """)

    def update_style(self):
        bg = THEME["accent"] if self.is_primary else THEME["bg_tertiary"]
        text = THEME["bg_primary"] if self.is_primary else THEME["text_primary"]
        border = THEME["accent"] if self.is_primary else THEME["border"]
        
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {bg};
                color: {text};
                border: 1px solid {border};
                border-radius: 4px;
                padding: 0 16px;
                font-weight: bold;
                font-family: 'Segoe UI', sans-serif;
            }}
            QPushButton:hover {{
                background-color: {THEME['accent_hover']} if {self.is_primary} else "{THEME['bg_secondary']}";
                border-color: {THEME['accent_hover']};
            }}
        """)

class ToggleSwitch(QWidget):
    toggled = pyqtSignal(bool)

    def __init__(self, checked=False, parent=None):
        super().__init__(parent)
        self.setFixedSize(50, 26)
        self.isChecked = checked
        self._circle_position = 22.0 if checked else 4.0
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.anim = QPropertyAnimation(self, b"circle_position", self)
        self.anim.setEasingCurve(QEasingCurve.Type.InOutQuad)
        self.anim.setDuration(150)

    @pyqtProperty(float)
    def circle_position(self):
        return self._circle_position

    @circle_position.setter
    def circle_position(self, pos):
        self._circle_position = pos
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        if self.isChecked:
            bg_color = QColor(THEME["accent"])
        else:
            bg_color = QColor(THEME["bg_tertiary"])
        
        painter.setBrush(bg_color)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(0, 0, self.width(), self.height(), 13, 13)

        painter.setBrush(QColor("#ffffff"))
        painter.drawEllipse(int(self._circle_position), 4, 18, 18)

    def mouseReleaseEvent(self, event):
        self.anim.stop()
        self.isChecked = not self.isChecked
        self.toggled.emit(self.isChecked)
        self.anim.setStartValue(self._circle_position)
        self.anim.setEndValue(22.0 if self.isChecked else 4.0)
        self.anim.start()

class ModCard(QFrame):
    status_changed = pyqtSignal(str, bool)
    delete_requested = pyqtSignal(str) # Signal for deletion request
    
    def __init__(self, mod_data):
        super().__init__()
        self.mod_data = mod_data
        self.setFixedSize(290, 95)
        self.is_delete_mode = False
        
        # Define styles
        self.default_style = f"""
            ModCard {{
                background-color: {THEME['bg_secondary']};
                border-radius: 8px;
                border: 1px solid {THEME['bg_tertiary']};
            }}
            ModCard:hover {{
                background-color: {THEME['bg_tertiary']};
                border: 1px solid {THEME['accent']};
            }}
        """
        
        self.delete_style = f"""
            ModCard {{
                background-color: {THEME['accent']};
                border-radius: 8px;
                border: 1px solid {THEME['accent']};
            }}
        """
        
        self.setStyleSheet(self.default_style)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setYOffset(4)
        shadow.setColor(QColor(0, 0, 0, 80))
        self.setGraphicsEffect(shadow)
        
        # Main Layout is now a Stack to swap between Normal and Delete views
        self.stack_layout = QStackedLayout(self)
        self.stack_layout.setContentsMargins(0, 0, 0, 0)

        # --- VIEW 1: Normal Content ---
        self.normal_view = QWidget()
        layout = QHBoxLayout(self.normal_view)
        layout.setContentsMargins(10, 5, 10, 5)
        layout.setSpacing(12)

        # Icon
        self.icon_lbl = QLabel()
        self.icon_lbl.setFixedSize(60, 60)
        
        has_image = False
        if "icon_data" in mod_data and mod_data["icon_data"]:
            pixmap = QPixmap()
            if pixmap.loadFromData(mod_data["icon_data"]):
                self.icon_lbl.setPixmap(pixmap)
                self.icon_lbl.setScaledContents(True)
                self.icon_lbl.setStyleSheet(f"border-radius: 8px; border: 1px solid {THEME['border']};")
                has_image = True
        
        if not has_image:
            self.icon_lbl.setText(mod_data["name"][:2].upper())
            self.icon_lbl.setStyleSheet(f"""
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {THEME['bg_tertiary']}, stop:1 {THEME['bg_primary']});
                border-radius: 8px;
                color: {THEME['text_primary']};
                font-size: 20px; font-weight: bold; border: 1px solid {THEME['border']};
            """)
        
        self.icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.icon_lbl)

        # Text Area
        text_layout = QVBoxLayout()
        text_layout.setSpacing(0)
        text_layout.setContentsMargins(0, 0, 0, 0)
        
        self.title = QLabel(mod_data["name"])
        self.title.setStyleSheet(f"font-size: 14px; font-weight: bold; color: {THEME['text_primary']}; background: transparent; border: none;")
        
        self.meta = QLabel(f"v{mod_data['version']} • {mod_data['author']}")
        self.meta.setStyleSheet(f"font-size: 11px; font-weight: 600; color: {THEME['accent']}; background: transparent; border: none;")
        
        self.desc = QLabel()
        self.desc.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.desc.setStyleSheet(f"font-size: 11px; color: {THEME['text_secondary']}; background: transparent; border: none;")
        
        available_width = 136 
        full_text = mod_data.get("description", "")
        max_lines = 3
        
        metrics = QFontMetrics(self.desc.font())
        elided_text = self.elide_multiline_text(metrics, full_text, available_width, max_lines)
        
        self.desc.setText(elided_text)
        if elided_text != full_text:
            self.desc.setToolTip(full_text)

        text_layout.addWidget(self.title)
        text_layout.addWidget(self.meta)
        text_layout.addWidget(self.desc)
        text_layout.addStretch() 
        layout.addLayout(text_layout)

        # Toggle 
        self.toggle = ToggleSwitch(checked=mod_data["enabled"])
        self.toggle.toggled.connect(self.on_toggle)
        layout.addWidget(self.toggle, 0, Qt.AlignmentFlag.AlignVCenter)

        # --- VIEW 2: Delete Confirmation ---
        self.delete_view = QWidget()
        self.delete_view.setStyleSheet("background: transparent;")
        del_layout = QHBoxLayout(self.delete_view)
        del_layout.setContentsMargins(0,0,0,0)
        del_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Trash Icon + Question Mark
        del_label = QLabel("🗑 ?") 
        del_label.setStyleSheet(f"""
            font-size: 40px; 
            font-weight: bold; 
            color: #ffffff;
            background: transparent;
            border: none;
        """)
        del_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        del_sub = QLabel("Click to Delete")
        del_sub.setStyleSheet("font-size: 12px; color: white; font-weight: bold; border: none;")
        del_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)

        v_del = QVBoxLayout()
        v_del.addStretch()
        v_del.addWidget(del_label, 0, Qt.AlignmentFlag.AlignCenter)
        v_del.addWidget(del_sub, 0, Qt.AlignmentFlag.AlignCenter)
        v_del.addStretch()
        
        del_layout.addLayout(v_del)

        # Add views to stack
        self.stack_layout.addWidget(self.normal_view)
        self.stack_layout.addWidget(self.delete_view)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.RightButton:
            # Toggle Delete Mode
            if self.is_delete_mode:
                self.reset_to_normal()
            else:
                self.enter_delete_mode()
        elif event.button() == Qt.MouseButton.LeftButton:
            # If in delete mode, confirm deletion
            if self.is_delete_mode:
                self.delete_requested.emit(self.mod_data["id"])
            else:
                # Pass through if not deleting
                super().mousePressEvent(event)
    
    def enterEvent(self, event):
        super().enterEvent(event)
        
    def leaveEvent(self, event):
        # Reset if mouse leaves the card
        if self.is_delete_mode:
            self.reset_to_normal()
        super().leaveEvent(event)

    def enter_delete_mode(self):
        self.is_delete_mode = True
        self.setStyleSheet(self.delete_style)
        self.stack_layout.setCurrentIndex(1)

    def reset_to_normal(self):
        self.is_delete_mode = False
        self.setStyleSheet(self.default_style)
        self.stack_layout.setCurrentIndex(0)

    def elide_multiline_text(self, metrics, text, width, max_lines):
        if metrics.horizontalAdvance(text) <= width:
            return text
        lines = []
        words = text.split(' ')
        current_line = []
        for word in words:
            test_line = ' '.join(current_line + [word])
            if metrics.horizontalAdvance(test_line) <= width:
                current_line.append(word)
            else:
                if current_line: lines.append(' '.join(current_line))
                current_line = [word]
        if current_line: lines.append(' '.join(current_line))
        if len(lines) <= max_lines: return "\n".join(lines)
        visible_lines = lines[:max_lines-1]
        remaining_text = " ".join(lines[max_lines-1:]) + " " + " ".join(lines[max_lines:])
        last_line = metrics.elidedText(remaining_text, Qt.TextElideMode.ElideRight, width)
        return "\n".join(visible_lines + [last_line])

    def on_toggle(self, checked):
        self.mod_data["enabled"] = checked
        self.status_changed.emit(self.mod_data["id"], checked)

class DashboardPage(QWidget):
    def __init__(self, mods_data):
        super().__init__()
        self.mods_data = mods_data
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(20)

        # Stats
        self.stats_layout = QHBoxLayout()
        self.stats_layout.setSpacing(20)
        layout.addLayout(self.stats_layout)
        
        self.refresh_stats()

        # Changelog
        changelog_frame = QFrame()
        changelog_frame.setStyleSheet(f"""
            background-color: {THEME['bg_secondary']};
            border-radius: 10px;
            border: 1px solid {THEME['bg_tertiary']};
        """)
        cl_layout = QVBoxLayout(changelog_frame)
        cl_layout.setContentsMargins(20, 20, 20, 20)
        
        cl_title = QLabel("LATEST UPDATES")
        cl_title.setStyleSheet(f"color: {THEME['text_secondary']}; font-size: 12px; font-weight: bold; letter-spacing: 1px; border: none;")
        cl_layout.addWidget(cl_title)
        cl_layout.addSpacing(10)

        # Updated Changelog entries (8 slots)
        all_updates = [
             ("v0.5.8", "UI Updates: 'Browse Online' placeholder and fixed monochrome icons."),
             ("v0.5.7", "Added 'Dev Mode' toggle to Settings (Controls console logs)."),
             ("v0.5.6", "Added 'Right-Click' delete mode to mod cards."),
             ("v0.5.5", "Implemented robust stateful XML parser (Fixes OptLib)."),
             ("v0.5.4", "Added custom logo support from /data/ folder."),
             ("v0.5.3", "Added 'Launch HoN' button with persistency."),
             ("v0.5.2", "Added support for asset injection mods."),
             ("v0.5.1", "Fixed .jz extraction (Zstd support)."),
        ]
        
        display_updates = all_updates[:8]

        for ver, desc in display_updates:
            row = QWidget()
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(0, 5, 0, 5)
            
            lbl_ver = QLabel(ver)
            lbl_ver.setFixedWidth(60)
            lbl_ver.setStyleSheet(f"color: {THEME['accent']}; font-weight: bold; border: none;")
            
            lbl_desc = QLabel(desc)
            lbl_desc.setStyleSheet(f"color: {THEME['text_primary']}; border: none;")
            lbl_desc.setWordWrap(True)
            
            row_layout.addWidget(lbl_ver)
            row_layout.addWidget(lbl_desc)
            cl_layout.addWidget(row)
            
            if ver != display_updates[-1][0]:
                line = QFrame()
                line.setFixedHeight(1)
                line.setStyleSheet(f"background-color: {THEME['bg_tertiary']}; border: none;")
                cl_layout.addWidget(line)
            
        cl_layout.addStretch()
        layout.addWidget(changelog_frame)

    def refresh_stats(self):
        while self.stats_layout.count():
            item = self.stats_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        total = len(self.mods_data)
        enabled = sum(1 for m in self.mods_data if m['enabled'])
        
        self.stats_layout.addWidget(self.create_stat_card("INSTALLED MODS", total, THEME['text_primary']))
        self.stats_layout.addWidget(self.create_stat_card("ACTIVE MODS", enabled, THEME['accent']))

    def create_stat_card(self, title, value, color):
        card = QFrame()
        card.setStyleSheet(f"""
            background-color: {THEME['bg_secondary']};
            border-radius: 10px;
            border: 1px solid {THEME['bg_tertiary']};
        """)
        card_layout = QVBoxLayout(card)
        
        lbl_title = QLabel(title)
        lbl_title.setStyleSheet(f"color: {THEME['text_secondary']}; font-size: 14px; font-weight: 600; border: none;")
        
        lbl_value = QLabel(str(value))
        lbl_value.setStyleSheet(f"color: {color}; font-size: 36px; font-weight: bold; border: none;")
        
        card_layout.addWidget(lbl_title)
        card_layout.addWidget(lbl_value)
        return card

class SettingsPage(QWidget):
    def __init__(self, dev_mode_enabled, toggle_callback):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.setSpacing(20)

        # Header
        header = QLabel("Application Settings")
        header.setStyleSheet(f"color: {THEME['text_primary']}; font-size: 18px; font-weight: bold;")
        layout.addWidget(header)

        # Developer Mode Option
        dev_frame = QFrame()
        dev_frame.setStyleSheet(f"""
            background-color: {THEME['bg_secondary']};
            border-radius: 8px;
            border: 1px solid {THEME['bg_tertiary']};
        """)
        dev_layout = QHBoxLayout(dev_frame)
        dev_layout.setContentsMargins(20, 20, 20, 20)

        text_layout = QVBoxLayout()
        lbl_title = QLabel("Developer Mode")
        lbl_title.setStyleSheet(f"color: {THEME['text_primary']}; font-size: 14px; font-weight: bold; border: none;")
        
        lbl_desc = QLabel("Enable console output for debugging mod application scripts.")
        lbl_desc.setStyleSheet(f"color: {THEME['text_secondary']}; font-size: 12px; border: none;")
        
        text_layout.addWidget(lbl_title)
        text_layout.addWidget(lbl_desc)

        self.dev_toggle = ToggleSwitch(checked=dev_mode_enabled)
        self.dev_toggle.toggled.connect(toggle_callback)

        dev_layout.addLayout(text_layout)
        dev_layout.addWidget(self.dev_toggle)

        layout.addWidget(dev_frame)
        layout.addStretch()

# --- LOGIC CLASS: MOD APPLICATOR ---
class ModApplicator:
    def __init__(self, game_root, resources_path, mods_data):
        self.game_root = Path(game_root)
        self.resources0_path = Path(resources_path) 
        self.mods_data = mods_data
        self.output_dir = self.game_root / "extensions"
        
        self.temp_output_path = self.output_dir / "resources0.zip"
        self.final_output_path = self.output_dir / "resources0.jz"

    def apply_mods(self):
        """Reads resources0, parses mods, and writes resources0.jz"""
        enabled_mods = [m for m in self.mods_data if m['enabled']]
        
        compression_method = zipfile.ZIP_DEFLATED
        if HAS_ZSTD:
            compression_method = getattr(zipfile, "ZIP_ZSTANDARD", 93)
            print("Using Zstandard (Method 93) compression.")
        else:
            print("Zstandard not found. Falling back to ZIP_DEFLATED.")

        if not enabled_mods:
            print("No mods enabled. Cleaning up resources0.jz...")
            if self.final_output_path.exists():
                os.remove(self.final_output_path)
            return True, "All mods disabled. (Cleaned up resources0.jz)"

        if not self.resources0_path.exists():
            return False, f"Base archive not found: {self.resources0_path}"

        print(f"--- STARTING APPLICATION: {len(enabled_mods)} Mods ---")
        
        modified_files = {} 

        try:
            for mod in enabled_mods:
                print(f"Processing: {mod['name']}")
                self._process_mod(mod, modified_files)

            if not self.output_dir.exists():
                self.output_dir.mkdir(parents=True)
            
            with zipfile.ZipFile(self.temp_output_path, 'w', compression_method) as z_out:
                for file_path, content in modified_files.items():
                    # print(f"Writing {file_path}...")
                    z_out.writestr(file_path, content)

            if self.final_output_path.exists():
                os.remove(self.final_output_path)
            os.rename(self.temp_output_path, self.final_output_path)
            
            return True, f"Successfully applied {len(enabled_mods)} mods to resources0.jz"

        except Exception as e:
            print(f"Error applying mods: {e}")
            if self.temp_output_path.exists():
                os.remove(self.temp_output_path)
            return False, str(e)

    def _process_mod(self, mod, modified_files):
        """Parses mod.xml using a STATEFUL cursor to handle multiple edits on the same file."""
        try:
            with zipfile.ZipFile(mod['file_path'], 'r') as z_mod:
                
                # 1. ASSET INJECTION
                ignored_files = {'mod.xml', 'icon.png', 'changelog.txt', 'icon.jpg', 'thumbs.db'}
                
                for file_info in z_mod.infolist():
                    fname = file_info.filename
                    fname_clean = fname.replace("\\", "/")
                    if file_info.is_dir() or fname_clean.lower() in ignored_files: continue
                    print(f"  [Asset] Injecting file: {fname_clean}")
                    modified_files[fname_clean] = z_mod.read(fname)

                # 2. XML PATCHING
                if 'mod.xml' in z_mod.namelist():
                    with z_mod.open('mod.xml') as f:
                        tree = ET.parse(f)
                        root = tree.getroot()

                        # Process each <editfile> block
                        for edit_node in root.findall(".//editfile"):
                            target_file = edit_node.get("name")
                            if not target_file: continue

                            # Load file content
                            if target_file not in modified_files:
                                self._load_base_file(target_file, modified_files)

                            current_content_bytes = modified_files.get(target_file)
                            if not current_content_bytes: continue

                            try:
                                # Decode
                                try:
                                    text_content = current_content_bytes.decode('utf-8')
                                except UnicodeDecodeError:
                                    text_content = current_content_bytes.decode('latin-1')

                                # Detect Line Endings of the target file
                                if "\r\n" in text_content:
                                    newline_mode = "\r\n"
                                else:
                                    newline_mode = "\n"

                                # --- STATEFUL CURSOR VARIABLES ---
                                cursor = 0
                                match_start = -1
                                match_end = -1
                                
                                # Iterate through operations sequentially
                                for child in edit_node:
                                    tag = child.tag.lower()
                                    
                                    # Normalize the search text to match the FILE's newline mode
                                    raw_text = child.text or ""
                                    search_text = raw_text.replace("\r\n", "\n").replace("\n", newline_mode)
                                    
                                    if tag == "find":
                                        # Search FORWARD from current cursor
                                        found_idx = text_content.find(search_text, cursor)
                                        if found_idx != -1:
                                            match_start = found_idx
                                            match_end = found_idx + len(search_text)
                                            # Update cursor to end of this match
                                            cursor = match_end
                                            print(f"  [Find] Match found at {match_start}")
                                        else:
                                            print(f"  [Miss] <find> failed for: {search_text[:20]}...")

                                    elif tag == "findup":
                                        # Check for position="end"
                                        start_pos = len(text_content)
                                        if child.get("position") != "end":
                                            start_pos = cursor

                                        # rfind searches backwards from start_pos
                                        found_idx = text_content.rfind(search_text, 0, start_pos)
                                        if found_idx != -1:
                                            match_start = found_idx
                                            match_end = found_idx + len(search_text)
                                            cursor = match_end 
                                            print(f"  [FindUp] Match found at {match_start}")
                                        else:
                                            print(f"  [Miss] <findup> failed for: {search_text[:20]}...")

                                    elif tag == "replace":
                                        if match_start != -1:
                                            # Replace the last matched text
                                            text_content = (
                                                text_content[:match_start] + 
                                                search_text + 
                                                text_content[match_end:]
                                            )
                                            # Update indices
                                            match_end = match_start + len(search_text)
                                            cursor = match_end
                                            print(f"  [Replace] Applied.")
                                        else:
                                            print("  [Skip] Replace called with no active match.")

                                    elif tag == "insert":
                                        if match_start != -1:
                                            position = child.get("position", "after")
                                            
                                            if position == "before":
                                                text_content = (
                                                    text_content[:match_start] + 
                                                    search_text + 
                                                    text_content[match_start:]
                                                )
                                                shift = len(search_text)
                                                match_start += shift
                                                match_end += shift
                                                cursor += shift
                                                print(f"  [Insert] Inserted Before.")

                                            else: # Default 'after'
                                                text_content = (
                                                    text_content[:match_end] + 
                                                    search_text + 
                                                    text_content[match_end:]
                                                )
                                                cursor += len(search_text)
                                                print(f"  [Insert] Inserted After.")
                                        else:
                                            print("  [Skip] Insert called with no active match.")

                                modified_files[target_file] = text_content.encode('utf-8')

                            except Exception as e:
                                print(f"  [Error] Failed processing edits for {target_file}: {e}")
        except Exception as e:
            print(f"Error processing mod {mod['name']}: {e}")

    def _load_base_file(self, filename, modified_files):
        if filename in modified_files: return
        try:
            with zipfile.ZipFile(self.resources0_path, 'r') as z_base:
                if filename in z_base.namelist():
                    modified_files[filename] = z_base.read(filename)
                else:
                    print(f"  [Error] File {filename} not found in resources0.jz")
        except Exception as e:
            print(f"  [Error] Reading base archive: {e}")


# --- MAIN WINDOW ---
class HoNModManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HoN Mod Manager")
        self.resize(936, 620)
        self.is_ready_to_launch = False 
        
        # Save original stdout for restoring later
        self.original_stdout = sys.stdout
        
        # --- ICON LOADING LOGIC ---
        # First try loading the ICO (preferred for windows), then the PNG
        icon_path_ico = resource_path(os.path.join("data", "HoNModManager.ico"))
        icon_path_png = resource_path(os.path.join("data", "HoNModManager.png"))
        
        app_icon = QIcon()
        if os.path.exists(icon_path_ico):
            app_icon.addFile(icon_path_ico)
        if os.path.exists(icon_path_png):
            app_icon.addFile(icon_path_png)
            
        if not app_icon.isNull():
            self.setWindowIcon(app_icon)
        
        if not HAS_ZSTD:
            QMessageBox.critical(self, "Missing Dependency", 
                "The library 'zipfile-zstd' is required to read game archives.\n\n"
                "Please run: pip install zipfile-zstd")

        self.game_exe_path = self.detect_game_path()
        self.game_root = None
        self.mods_dir = None
        self.resources_path = None 

        if self.game_exe_path:
            self.game_root = Path(self.game_exe_path).parent.parent
            self.prepare_game_folders(str(self.game_root))
            self.detect_resources_file() 
        
        if self.game_root:
            res1 = self.game_root / "extensions" / "resources0.jz"
            if res1.exists():
                self.is_ready_to_launch = True

        # Load Config (Mods + Dev Mode)
        self.saved_enabled_mods, self.dev_mode = self.load_config()
        
        # Apply Dev Mode Setting Immediately
        self.update_dev_mode(self.dev_mode)

        self.refresh_mods_library()

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QHBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        self.setup_ui()
        self.apply_styles()
        self.update_action_button()

    def load_config(self):
        enabled_mods = set()
        dev_mode = False # Default to OFF
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    data = json.load(f)
                    enabled_mods = set(data.get("enabled_mods", []))
                    dev_mode = data.get("dev_mode", False)
            except:
                pass
        return enabled_mods, dev_mode

    def save_config(self):
        enabled = [m["id"] for m in INSTALLED_MODS if m["enabled"]]
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump({
                    "enabled_mods": enabled,
                    "dev_mode": self.dev_mode
                }, f)
        except Exception as e:
            # We can't print if dev mode is off, but this is a critical error
            # so we force it to original stdout momentarily
            sys.stdout = self.original_stdout
            print(f"Error saving config: {e}")
            if not self.dev_mode: sys.stdout = NullWriter()

    def update_dev_mode(self, enabled):
        self.dev_mode = enabled
        if enabled:
            sys.stdout = self.original_stdout
            print("Dev Mode Enabled: Console output restored.")
        else:
            print("Dev Mode Disabled: Console output suppressed.")
            sys.stdout = NullWriter()
        self.save_config()

    def detect_game_path(self):
        print("--- Scanning Registry for Juvio ---")
        reg_paths = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Uninstall")
        ]
        for hive, path in reg_paths:
            try:
                with winreg.OpenKey(hive, path) as key:
                    for i in range(winreg.QueryInfoKey(key)[0]):
                        try:
                            with winreg.OpenKey(key, winreg.EnumKey(key, i)) as subkey:
                                try:
                                    display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                except OSError:
                                    continue
                                if "juvio" in display_name.lower():
                                    try:
                                        root = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                                    except FileNotFoundError:
                                        u_str = winreg.QueryValueEx(subkey, "UninstallString")[0]
                                        root = os.path.dirname(u_str.split('"')[1] if '"' in u_str else u_str.split()[0])
                                    exe_path = os.path.join(root, "bin", "juvio.exe")
                                    if os.path.exists(exe_path):
                                        return exe_path
                                    else:
                                        root_exe = os.path.join(root, "juvio.exe")
                                        if os.path.exists(root_exe):
                                            return root_exe 
                        except (OSError, IndexError): continue
            except OSError: continue
        return None

    def prepare_game_folders(self, game_dir_str):
        try:
            game_dir = Path(game_dir_str)
            ext_dir = game_dir / "extensions"
            self.mods_dir = ext_dir / "mods"
            if not ext_dir.exists(): ext_dir.mkdir()
            if not self.mods_dir.exists(): self.mods_dir.mkdir()
        except Exception as e:
            print(f"CRITICAL ERROR preparing folders: {e}")

    def detect_resources_file(self):
        if not self.game_root: return
        target_path = self.game_root / "heroes of newerth" / "resources0.jz"
        if target_path.exists():
            self.resources_path = target_path
        else:
            print(f"WARNING: resources0.jz NOT found at {target_path}")

    def parse_mod_file(self, file_path):
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                if 'mod.xml' not in z.namelist(): return None
                with z.open('mod.xml') as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    def get_meta(key, default="Unknown"):
                        val = root.get(key)
                        if val is None: val = root.findtext(key)
                        return val if val else default
                    name = get_meta('name', file_path.stem)
                    version = get_meta('version', "1.0")
                    author = get_meta('author', "Unknown")
                    description = get_meta('description', "No description provided.")
                    mod_id = name.lower().replace(" ", "_")
                icon_path = root.get('icon', 'icon.png')
                icon_data = None
                if icon_path in z.namelist(): icon_data = z.read(icon_path)
                elif 'icon.png' in z.namelist(): icon_data = z.read('icon.png')
                is_enabled = mod_id in self.saved_enabled_mods
                return {
                    "id": mod_id, "name": name, "author": author, "version": version, 
                    "category": "Installed", "enabled": is_enabled, "description": description, 
                    "icon_data": icon_data, "file_path": str(file_path)
                }
        except Exception as e:
            print(f"Error parsing {file_path}: {e}")
            return None

    def refresh_mods_library(self):
        if not self.mods_dir or not self.mods_dir.exists(): return
        INSTALLED_MODS.clear()
        files = list(self.mods_dir.glob("*.honmod")) + list(self.mods_dir.glob("*.zip"))
        for file_path in files:
            mod_data = self.parse_mod_file(file_path)
            if mod_data: INSTALLED_MODS.append(mod_data)

    def import_mods(self):
        if not self.mods_dir: return
        file_names, _ = QFileDialog.getOpenFileNames(self, "Import Mods", "", "HoN Mods (*.honmod);;Zip Files (*.zip)")
        if not file_names: return
        for src_path in file_names:
            try:
                src = Path(src_path)
                dst = self.mods_dir / src.name
                shutil.copy2(src, dst)
            except Exception as e:
                print(f"Error importing {src_path}: {e}")
        self.refresh_mods_library()
        self.populate_mods()
        self.dashboard_page.refresh_stats()

    def update_action_button(self):
        self.apply_btn.set_launch_mode(self.is_ready_to_launch)

    def on_main_button_click(self):
        if self.is_ready_to_launch: self.launch_game()
        else: self.apply_mods()

    def launch_game(self):
        if not self.game_exe_path:
            QMessageBox.critical(self, "Error", "Game executable not found.")
            return
        try:
            cmd = [self.game_exe_path, "-mod", "heroes of newerth;extensions"]
            print(f"Launching: {cmd}")
            subprocess.Popen(cmd, cwd=self.game_root)
        except Exception as e:
            QMessageBox.critical(self, "Launch Error", str(e))

    def apply_mods(self):
        if not self.game_root or not self.resources_path:
            QMessageBox.critical(self, "Error", "Game resources not found. Cannot apply mods.")
            return
        if not HAS_ZSTD:
            QMessageBox.critical(self, "Missing Library", "Cannot apply mods: 'zipfile-zstd' is missing.")
            return
        self.apply_btn.setText("Applying...")
        self.apply_btn.setEnabled(False)
        QApplication.processEvents()
        applicator = ModApplicator(self.game_root, self.resources_path, INSTALLED_MODS)
        success, msg = applicator.apply_mods()
        self.apply_btn.setEnabled(True)
        if success:
            self.is_ready_to_launch = True
            # Config is saved by handle_toggle, but good to be safe
            self.update_action_button()
            QMessageBox.information(self, "Success", msg)
        else:
            self.is_ready_to_launch = False
            self.update_action_button()
            QMessageBox.warning(self, "Mod Application", msg)

    def delete_mod(self, mod_id):
        # Find mod in list
        mod_to_delete = next((m for m in INSTALLED_MODS if m['id'] == mod_id), None)
        if not mod_to_delete: return
        
        file_path = Path(mod_to_delete['file_path'])
        mod_name = mod_to_delete['name']
        
        try:
            # Delete file
            if file_path.exists():
                os.remove(file_path)
                print(f"Deleted {file_path}")
            
            # Update memory
            INSTALLED_MODS.remove(mod_to_delete)
            
            # Update config
            if mod_id in self.saved_enabled_mods:
                self.saved_enabled_mods.remove(mod_id)
                self.save_config()
            
            # Refresh UI
            self.populate_mods()
            self.dashboard_page.refresh_stats()
            
            # Reset Launch State
            self.is_ready_to_launch = False
            self.update_action_button()
            
        except Exception as e:
            QMessageBox.critical(self, "Deletion Error", f"Could not delete {mod_name}:\n{str(e)}")

    def populate_mods(self):
        while self.mod_layout.count():
            item = self.mod_layout.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self.mod_cards.clear()
        for mod in INSTALLED_MODS:
            card = ModCard(mod)
            card.status_changed.connect(self.handle_toggle)
            card.delete_requested.connect(self.delete_mod)
            self.mod_layout.addWidget(card)
            self.mod_cards.append(card)

    def setup_ui(self):
        self.sidebar = QFrame()
        self.sidebar.setFixedWidth(260)
        self.sidebar_layout = QVBoxLayout(self.sidebar)
        self.sidebar_layout.setContentsMargins(0, 0, 0, 20)
        self.sidebar_layout.setSpacing(2) 
        logo_area = QFrame()
        logo_area.setFixedHeight(120)
        logo_layout = QHBoxLayout(logo_area)
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        img_path = Path(__file__).parent / "data" / "HoNModManager.png"
        if img_path.exists():
            pix = QPixmap(str(img_path))
            pix = pix.scaled(220, 100, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(pix)
            logo_label.setStyleSheet("border: none;")
        else:
            logo_label.setText("HON MODS")
            logo_label.setStyleSheet(f"font-family: 'Impact'; font-size: 32px; letter-spacing: 2px; color: {THEME['text_primary']}; border: none;")
        logo_layout.addWidget(logo_label)
        self.sidebar_layout.addWidget(logo_area)
        
        self.nav_btns = []
        self.nav_btns.append(SidebarButton("Dashboard", "☷"))
        # CHANGED: "My Mods" icon -> "Squared Times" (⊠)
        # This is a standard math symbol that looks like a crate but renders in monochrome.
        self.nav_btns.append(SidebarButton("My Mods", "⊠"))
        # CHANGED: Added is_locked=True for the Online tab
        self.nav_btns.append(SidebarButton("Browse Online", "☁", is_locked=True))
        self.nav_btns.append(SidebarButton("Settings", "⚙"))
        
        for btn in self.nav_btns:
            if not btn.is_locked:
                btn.clicked.connect(lambda checked, b=btn: self.on_nav_click(b))
            self.sidebar_layout.addWidget(btn)
            
        self.sidebar_layout.addStretch()
        
        # --- FOOTER UPDATE ---
        footer_label = QLabel("HoN Mod Manager v0.5.8\nBy Wartype")
        footer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer_label.setStyleSheet(f"color: {THEME['text_secondary']}; font-size: 11px; border: none;")
        self.sidebar_layout.addWidget(footer_label)
        
        self.main_layout.addWidget(self.sidebar)
        self.content_area = QFrame()
        self.content_layout = QVBoxLayout(self.content_area)
        self.content_layout.setContentsMargins(40, 40, 40, 40)
        header_layout = QHBoxLayout()
        self.page_title = QLabel("Dashboard")
        self.page_title.setStyleSheet(f"font-size: 32px; font-weight: 800; color: {THEME['text_primary']}; border: none;")
        display_path = str(self.game_root) if self.game_root else "Game Not Found"
        self.path_label = QLabel(display_path)
        self.path_label.setStyleSheet(f"color: {THEME['text_secondary']}; font-family: Consolas; border: none;")
        header_layout.addWidget(self.page_title)
        header_layout.addStretch()
        header_layout.addWidget(self.path_label)
        self.content_layout.addLayout(header_layout)
        self.content_layout.addSpacing(20)
        self.stack = QStackedWidget()
        self.content_layout.addWidget(self.stack)
        
        # Page 0: Dashboard
        self.dashboard_page = DashboardPage(INSTALLED_MODS)
        self.stack.addWidget(self.dashboard_page)
        
        # Page 1: My Mods
        self.mods_page = QWidget()
        mods_layout = QVBoxLayout(self.mods_page)
        mods_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout = QHBoxLayout()
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search installed mods...")
        self.search_bar.setFixedHeight(40)
        self.search_bar.textChanged.connect(self.filter_mods)
        self.import_btn = ModernButton("Import Mods", is_primary=False)
        self.import_btn.setFixedHeight(40)
        self.import_btn.clicked.connect(self.import_mods)
        self.apply_btn = ModernButton("Apply Changes", is_primary=True)
        self.apply_btn.setFixedHeight(40)
        self.apply_btn.clicked.connect(self.on_main_button_click)
        controls_layout.addWidget(self.search_bar, 1)
        controls_layout.addWidget(self.import_btn)
        controls_layout.addSpacing(10)
        controls_layout.addWidget(self.apply_btn)
        mods_layout.addLayout(controls_layout)
        mods_layout.addSpacing(20)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.mod_container = QWidget()
        self.mod_container.setObjectName("mod_container")
        self.mod_layout = FlowLayout(self.mod_container, margin=0, h_spacing=15, v_spacing=15)
        self.mod_cards = []
        self.populate_mods()
        self.scroll.setWidget(self.mod_container)
        mods_layout.addWidget(self.scroll)
        self.stack.addWidget(self.mods_page)

        # Page 2: Online (Placeholder)
        online_page = QWidget()
        ol_layout = QVBoxLayout(online_page)
        ol_lbl = QLabel("Online Mod Browser Coming Soon")
        ol_lbl.setStyleSheet("color: #666; font-size: 24px;")
        ol_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ol_layout.addWidget(ol_lbl)
        self.stack.addWidget(online_page)

        # Page 3: Settings
        self.settings_page = SettingsPage(self.dev_mode, self.update_dev_mode)
        self.stack.addWidget(self.settings_page)

        self.main_layout.addWidget(self.content_area)
        self.nav_btns[0].setChecked(True)
        self.stack.setCurrentIndex(0)

    def filter_mods(self, text):
        text = text.lower()
        for card in self.mod_cards:
            match = text in card.mod_data["name"].lower()
            card.setVisible(match)

    def on_nav_click(self, clicked_btn):
        for btn in self.nav_btns: btn.setChecked(btn == clicked_btn)
        page_name = clicked_btn.text_lbl.text()
        self.page_title.setText(page_name)
        if page_name == "Dashboard": self.stack.setCurrentIndex(0)
        elif page_name == "My Mods": self.stack.setCurrentIndex(1)
        elif page_name == "Browse Online": self.stack.setCurrentIndex(2)
        elif page_name == "Settings": self.stack.setCurrentIndex(3)
        
    def handle_toggle(self, mod_id, status):
        self.is_ready_to_launch = False
        self.update_action_button()
        self.dashboard_page.refresh_stats()
        # Save mod status to config
        self.save_config()

    def apply_styles(self):
        self.setStyleSheet(f"""
            QMainWindow {{ background-color: {THEME['bg_primary']}; }}
            QWidget {{ font-family: 'Segoe UI', sans-serif; }}
            QFrame {{ border: none; }}
            QFrame#sidebar {{ background-color: {THEME['bg_secondary']}; border-right: 1px solid {THEME['border']}; }}
            QWidget#mod_container {{ background-color: {THEME['bg_primary']}; }}
            QToolTip {{ color: {THEME['text_primary']}; background-color: {THEME['bg_tertiary']}; border: 1px solid {THEME['accent']}; border-radius: 4px; padding: 5px; }}
            QLineEdit {{ background-color: {THEME['bg_tertiary']}; color: {THEME['text_primary']}; border: 1px solid {THEME['border']}; border-radius: 6px; padding: 5px 15px; font-size: 14px; }}
            QLineEdit:focus {{ border: 1px solid {THEME['accent']}; }}
            QScrollBar:vertical {{ border: none; background: {THEME['bg_primary']}; width: 8px; }}
            QScrollBar::handle:vertical {{ background: {THEME['border']}; border-radius: 4px; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        """)
        self.sidebar.setObjectName("sidebar")

if __name__ == "__main__":
    # --- FIX FOR TASKBAR ICON ---
    # This block forces Windows to treat the script/exe as a distinct application
    # with its own icon in the taskbar.
    if sys.platform == 'win32':
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)

    app = QApplication(sys.argv)
    window = HoNModManager()
    window.show()
    sys.exit(app.exec())
