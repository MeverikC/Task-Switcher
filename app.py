import sys
import os
import json
import math
import win32gui
import win32con
import time
import win32api
import win32process
import ctypes
from ctypes import wintypes
import keyboard
import psutil
from PyQt6.QtWidgets import (QApplication, QListWidget, QListWidgetItem,
                             QVBoxLayout, QWidget, QStyle, QDialog, QFormLayout,
                             QSystemTrayIcon, QMenu, QStyledItemDelegate, QFileIconProvider,
                             QPushButton, QColorDialog, QSlider, QSpinBox, QRadioButton, QButtonGroup, QHBoxLayout,
                             QLabel, QFrame, QStyleOption)
from PyQt6.QtCore import Qt, QSize, pyqtSignal, QFileInfo, QRect, QTimer
from PyQt6.QtGui import QIcon, QAction, QColor, QPainter


user32 = ctypes.windll.user32
# DWM API
dwmapi = ctypes.WinDLL("dwmapi")
DWMWA_CLOAKED = 14

# SystemParametersInfo 常量，用于解决切换焦点时的 LockTimeout 问题
# SPI_GETFOREGROUNDLOCKTIMEOUT = 0x2000
# SPI_SETFOREGROUNDLOCKTIMEOUT = 0x2001


def is_window_cloaked(hwnd):
    is_cloaked = ctypes.c_int(0)
    try:
        hr = dwmapi.DwmGetWindowAttribute(hwnd, DWMWA_CLOAKED, ctypes.byref(is_cloaked), ctypes.sizeof(is_cloaked))
        if hr == 0:
            return is_cloaked.value != 0
    except:
        pass
    return False

def get_recourse_path(filename):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)

class ConfigManager:
    def __init__(self, filename="settings.json"):
        self.filename = filename
        self.default_settings = {
            "bg_color": "#edf2fa",
            "text_color": "#333333",
            "sel_bg_color": "#cce8ff",
            "opacity": 1.0,
            "layout_mode": "grid",
            "max_items": 6
        }
        self.settings = self.load_settings()

    def load_settings(self):
        if os.path.exists(self.filename):
            try:
                with open(self.filename, 'r') as f:
                    return {**self.default_settings, **json.load(f)}
            except:
                pass
        return self.default_settings.copy()

    def save_settings(self):
        with open(self.filename, 'w') as f:
            json.dump(self.settings, f, indent=4)

    def get(self, key):
        return self.settings.get(key)

    def set(self, key, value):
        self.settings[key] = value
        self.save_settings()


CONFIG = ConfigManager()


# ==========================================
# 2. 强化的自定义委托
# ==========================================
class UniversalDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.mode = CONFIG.get("layout_mode")

    def update_mode(self):
        self.mode = CONFIG.get("layout_mode")

    def paint(self, painter, option, index):
        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # 绘制背景（圆角）
        if option.state & QStyle.StateFlag.State_Selected:
            bg_color = QColor(CONFIG.get("sel_bg_color"))
            painter.setBrush(bg_color)
            painter.setPen(Qt.PenStyle.NoPen)
            # 留一点 margin 使得选中态更好看
            rect = option.rect.adjusted(2, 2, -2, -2)
            painter.drawRoundedRect(rect, 6, 6)

        # 获取数据
        icon = index.data(Qt.ItemDataRole.DecorationRole)
        text = index.data(Qt.ItemDataRole.DisplayRole)

        text_color = QColor(CONFIG.get("text_color"))
        painter.setPen(text_color)
        font = painter.font()
        font.setFamily("Microsoft YaHei UI")

        if self.mode == "list":
            icon_size = 32
            padding = 15

            # 垂直居中
            cy = option.rect.top() + option.rect.height() // 2

            # 图标
            icon_rect = QRect(option.rect.left() + padding, cy - icon_size // 2, icon_size, icon_size)
            if icon:
                icon.paint(painter, icon_rect)

            # 文字
            text_rect = QRect(icon_rect.right() + padding, option.rect.top(),
                              option.rect.width() - icon_rect.right() - padding * 2, option.rect.height())

            font.setPixelSize(14)
            painter.setFont(font)
            fm = painter.fontMetrics()
            elided_text = fm.elidedText(text, Qt.TextElideMode.ElideRight, text_rect.width())
            painter.drawText(text_rect, Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, elided_text)

        else:  # grid
            icon_size = 48
            # 图标居中，稍微偏上
            icon_x = option.rect.left() + (option.rect.width() - icon_size) // 2
            icon_y = option.rect.top() + 15
            icon_rect = QRect(int(icon_x), int(icon_y), icon_size, icon_size)

            if icon:
                icon.paint(painter, icon_rect)

            # 文字在下方
            text_rect = QRect(option.rect.left() + 4, icon_rect.bottom() + 8,
                              option.rect.width() - 8, 20)

            font.setPixelSize(12)
            painter.setFont(font)
            fm = painter.fontMetrics()
            elided_text = fm.elidedText(text, Qt.TextElideMode.ElideRight, text_rect.width())
            painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, elided_text)

        painter.restore()


# ==========================================
# 3. 现代化的设置窗口
# ==========================================
class SettingsDialog(QDialog):
    settings_changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Task Switcher Settings")
        self.resize(380, 450)
        self.setWindowIcon(QIcon(get_recourse_path("icon.png")))

        # 应用样式表：去除SpinBox箭头，美化输入框，增加间距
        self.setStyleSheet("""
            QDialog {
                background-color: #f9f9f9;
                font-family: "Microsoft YaHei UI";
                font-size: 14px;
            }
            QLabel {
                color: #555555;
                font-weight: bold;
                padding-top: 5px;
            }
            QPushButton.color-btn {
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 5px;
                text-align: left;
                padding-left: 10px;
                font-family: "Consolas", monospace;
            }
            QPushButton.color-btn:hover {
                border: 1px solid #888888;
            }
            QSpinBox {
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 5px;
                background: white;
            }
            QSpinBox::up-button, QSpinBox::down-button {
                width: 0px; 
                height: 0px;
                border: none; 
            }
            QRadioButton {
                padding: 5px;
                border: 1px solid transparent;
                border-radius: 4px;
            }
            QRadioButton::indicator {
                width: 16px;
                height: 16px;
            }
            QRadioButton:checked {
                background-color: #e6f7ff;
                border: 1px solid #1890ff;
                color: #096dd9;
            }
            QSlider::groove:horizontal {
                border: 1px solid #bbb;
                background: white;
                height: 6px;
                border-radius: 3px;
            }
            QSlider::sub-page:horizontal {
                background: #1890ff;
                border-radius: 3px;
            }
            QSlider::handle:horizontal {
                background: white;
                border: 1px solid #555;
                width: 16px;
                height: 16px;
                margin: -6px 0; 
                border-radius: 8px;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

        # 表单布局
        form_layout = QFormLayout()
        form_layout.setSpacing(15)  # 增加行间距
        form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)

        # 1. 外观颜色
        self.btn_bg = self.create_color_btn("bg_color")
        form_layout.addRow("窗口背景:", self.btn_bg)

        self.btn_text = self.create_color_btn("text_color")
        form_layout.addRow("文字颜色:", self.btn_text)

        self.btn_sel = self.create_color_btn("sel_bg_color")
        form_layout.addRow("选中背景:", self.btn_sel)

        # 2. 透明度
        self.slider_opacity = QSlider(Qt.Orientation.Horizontal)
        self.slider_opacity.setRange(20, 100)
        self.slider_opacity.setValue(int(CONFIG.get("opacity") * 100))
        self.slider_opacity.valueChanged.connect(self.update_opacity)
        form_layout.addRow("不透明度:", self.slider_opacity)

        # 3. 布局模式
        mode_container = QWidget()
        mode_layout = QHBoxLayout(mode_container)
        mode_layout.setContentsMargins(0, 0, 0, 0)

        self.radio_list = QRadioButton("列表模式")
        self.radio_grid = QRadioButton("网格模式")

        self.bg_group = QButtonGroup()
        self.bg_group.addButton(self.radio_list, 0)
        self.bg_group.addButton(self.radio_grid, 1)

        if CONFIG.get("layout_mode") == "list":
            self.radio_list.setChecked(True)
        else:
            self.radio_grid.setChecked(True)

        self.bg_group.idToggled.connect(self.update_layout_mode)

        mode_layout.addWidget(self.radio_list)
        mode_layout.addWidget(self.radio_grid)
        form_layout.addRow("布局方式:", mode_container)

        # 4. 数量限制
        self.spin_max = QSpinBox()
        self.spin_max.setRange(1, 50)
        self.spin_max.setValue(CONFIG.get("max_items"))
        self.spin_max.valueChanged.connect(lambda v: self.save_val("max_items", v))
        self.spin_max.setToolTip("列表模式为最大行数，网格模式为每行个数")
        form_layout.addRow("显示阈值:", self.spin_max)

        layout.addLayout(form_layout)

        # 底部说明
        note = QLabel("提示: 修改后即时生效，Alt+Tab 预览")
        note.setStyleSheet("color: #999; font-size: 12px; margin-top: 10px;")
        note.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(note)

    def create_color_btn(self, key):
        """创建带背景色和Hex文字的按钮"""
        hex_val = CONFIG.get(key)
        btn = QPushButton(hex_val)
        btn.setProperty("key", key)  # 存储 key
        btn.setProperty("class", "color-btn")
        self.update_btn_style(btn, hex_val)
        btn.clicked.connect(lambda: self.pick_color(key, btn))
        return btn

    def update_btn_style(self, btn, hex_color):
        # 计算亮度以决定文字颜色
        c = QColor(hex_color)
        text_color = "black" if c.lightness() > 128 else "white"
        btn.setText(hex_color.upper())
        # 利用 border-left 显示颜色块，或者直接背景
        btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {hex_color};
                color: {text_color};
            }}
        """)

    def pick_color(self, key, btn):
        color = QColorDialog.getColor(QColor(CONFIG.get(key)), self)
        if color.isValid():
            hex_color = color.name()
            CONFIG.set(key, hex_color)
            self.update_btn_style(btn, hex_color)
            self.settings_changed.emit()

    def update_opacity(self, value):
        CONFIG.set("opacity", value / 100.0)
        self.settings_changed.emit()

    def update_layout_mode(self, btn_id, checked):
        if checked:
            mode = "list" if btn_id == 0 else "grid"
            CONFIG.set("layout_mode", mode)
            self.settings_changed.emit()

    def save_val(self, key, val):
        CONFIG.set(key, val)
        self.settings_changed.emit()


# ==========================================
# 4. 主窗口
# ==========================================
class WindowSwitcher(QWidget):
    sig_show = pyqtSignal()
    sig_next = pyqtSignal()
    sig_activate = pyqtSignal()

    def __init__(self):
        super().__init__()

        # Tool 属性确保不显示在任务栏
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        self.icon_cache = {}
        self.icon_provider = QFileIconProvider()


        self.sig_show.connect(self.show_switcher)
        self.sig_next.connect(self.select_next)
        self.sig_activate.connect(self.activate_selected)

        # 1. 开启心跳定时器 (防止进程被系统挂起)
        self.heartbeat_timer = QTimer(self)
        self.heartbeat_timer.timeout.connect(self.on_heartbeat)
        self.heartbeat_timer.start(1000)  # 每 1 秒跳动一次

        # 2. 开启钩子守护定时器 (定期重置钩子，防止 Windows 杀钩子)
        self.hook_guard_timer = QTimer(self)
        self.hook_guard_timer.timeout.connect(self.reset_hooks)
        self.hook_guard_timer.start(1000 * 60 * 30)  # 每 30 分钟重置一次

        # 初始化钩子
        self.setup_hooks()

        self.init_ui()
        self.init_tray_icon()
        self.apply_settings()

    def init_ui(self):
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)

        self.list_widget = QListWidget()
        self.list_widget.setFrameShape(QListWidget.Shape.NoFrame)

        # --- 修改点：彻底关闭滚动条 ---
        self.list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.list_widget.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        # ---------------------------

        self.delegate = UniversalDelegate()
        self.list_widget.setItemDelegate(self.delegate)
        self.list_widget.itemClicked.connect(lambda item: self.activate_selected())

        self.layout.addWidget(self.list_widget)

    def on_heartbeat(self):
        """
        空操作，仅仅为了让 Qt 事件循环保持活跃，
        防止 Windows 认为进程空闲而将其挂起/降低优先级。
        """
        pass

    def setup_hooks(self):
        """挂载键盘钩子"""
        try:
            # 先卸载所有旧钩子，防止重复
            keyboard.unhook_all()

            # 重新挂载
            keyboard.add_hotkey('alt+tab', self.on_hotkey_tab, suppress=True)
            keyboard.on_release_key('alt', self.on_hotkey_release)
            print(f"Hooks installed/refreshed at {time.strftime('%H:%M:%S')}")
        except Exception as e:
            print(f"Hook Error: {e}")

    def reset_hooks(self):
        """定期重置钩子"""
        # print("Watchdog: Refreshing hooks...")
        self.setup_hooks()

    def on_hotkey_tab(self):
        if not self.isVisible():
            self.sig_show.emit()
        else:
            self.sig_next.emit()

    def on_hotkey_release(self, e):
        if self.isVisible():
            self.sig_activate.emit()

    def apply_settings(self):
        bg = CONFIG.get("bg_color")
        self.setObjectName("SwitcherMain")
        self.setStyleSheet(f"""
            #SwitcherMain {{
                background-color: {bg};
                border: 1px solid #888888;
                border-radius: 12px;  /* 这里修改圆角大小 */
            }}
            QListWidget {{
                background-color: transparent; /* 列表背景透明，透出父窗口颜色 */
                border: none;
                outline: none;
            }}
        """)
        self.setWindowOpacity(CONFIG.get("opacity"))
        self.delegate.update_mode()

        mode = CONFIG.get("layout_mode")

        # 统一设置边距
        self.layout.setContentsMargins(10, 10, 10, 10)

        if mode == "grid":
            self.list_widget.setViewMode(QListWidget.ViewMode.IconMode)
            self.list_widget.setFlow(QListWidget.Flow.LeftToRight)
            self.list_widget.setWrapping(True)  # 必须允许换行
            self.list_widget.setResizeMode(QListWidget.ResizeMode.Adjust)
            self.list_widget.setMovement(QListWidget.Movement.Static)  # 禁止拖拽移动，固定位置

            # 这里的 spacing 必须是 8，与上面计算一致
            self.list_widget.setSpacing(8)

        else:
            self.list_widget.setViewMode(QListWidget.ViewMode.ListMode)
            self.list_widget.setFlow(QListWidget.Flow.TopToBottom)
            self.list_widget.setSpacing(2)

        self.list_widget.update()

    def paintEvent(self, event):
        """
        核心修复：在开启透明背景属性后，必须手动绘制 QSS 样式，
        否则窗口就是完全透明不可见的。
        """
        opt = QStyleOption()
        opt.initFrom(self)
        p = QPainter(self)
        # 开启抗锯齿，让圆角平滑
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        # 使用 QSS 定义的样式（背景色、边框、圆角）进行绘制
        self.style().drawPrimitive(QStyle.PrimitiveElement.PE_Widget, opt, p, self)

    def open_settings(self):
        self.settings_dlg = SettingsDialog(self)
        self.settings_dlg.setWindowFlags(Qt.WindowType.Window)
        self.settings_dlg.settings_changed.connect(self.apply_settings)
        self.settings_dlg.show()

    def init_tray_icon(self):
        self.tray_icon = QSystemTrayIcon(self)

        icon_path = get_recourse_path("icon.png")
        if os.path.exists(icon_path):
            self.tray_icon.setIcon(QIcon(icon_path))
        else:
            self.tray_icon.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon))

        menu = QMenu()
        action_settings = QAction("设置...", self)
        action_settings.triggered.connect(self.open_settings)
        menu.addAction(action_settings)
        menu.addSeparator()
        action_quit = QAction("退出", self)
        action_quit.triggered.connect(self.quit_app)
        menu.addAction(action_quit)
        self.tray_icon.setContextMenu(menu)
        self.tray_icon.show()

    def quit_app(self):
        self.tray_icon.hide()
        QApplication.quit()

    def get_window_icon(self, hwnd):
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid in self.icon_cache: return self.icon_cache[pid]
            process = psutil.Process(pid)
            exe_path = process.exe()
            if os.path.exists(exe_path):
                icon = self.icon_provider.icon(QFileInfo(exe_path))
                self.icon_cache[pid] = icon
                return icon
        except:
            pass
        return self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon)

    def refresh_windows(self):
        self.list_widget.clear()

        def enum_handler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd) and not is_window_cloaked(hwnd):
                title = win32gui.GetWindowText(hwnd)
                # 过滤掉一些特定窗口
                if title and title != "Program Manager":
                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                    if not (style & win32con.WS_EX_TOOLWINDOW) or (style & win32con.WS_EX_APPWINDOW):
                        if int(hwnd) != int(self.winId()):
                            if hasattr(self, 'settings_dlg') and self.settings_dlg.isVisible():
                                if int(hwnd) == int(self.settings_dlg.winId()): return
                            self.add_window_item(hwnd, title)

        win32gui.EnumWindows(enum_handler, None)
        self.adjust_window_size()

    def add_window_item(self, hwnd, title):
        item = QListWidgetItem()
        item.setData(Qt.ItemDataRole.DisplayRole, title)
        item.setData(Qt.ItemDataRole.DecorationRole, self.get_window_icon(hwnd))
        item.setData(Qt.ItemDataRole.UserRole, hwnd)
        item.setToolTip(title)

        if CONFIG.get("layout_mode") == "list":
            item.setSizeHint(QSize(200, 60))
        else:
            # 必须是 110，与计算公式一致
            item.setSizeHint(QSize(110, 110))

        self.list_widget.addItem(item)

    def adjust_window_size(self):
        """
        修复版：
        1. 宽度增加余量，防止第6个被挤下去。
        2. 高度根据总行数完全撑开，不限制最大行数。
        """
        count = self.list_widget.count()
        if count == 0: return

        max_items = CONFIG.get("max_items")
        layout_mode = CONFIG.get("layout_mode")

        # 获取边距 (假设我们在 apply_settings 里设置了 margin)
        m_left, m_top, m_right, m_bottom = 10, 10, 10, 10

        if layout_mode == "list":
            # --- 列表模式 ---
            item_height = 60  # item hint
            spacing = 2

            # 高度 = 数量 * (高度 + 间距) + 上下边距
            total_h = count * (item_height + spacing) + m_top + m_bottom
            self.resize(360, total_h)

        else:
            # --- 网格模式 (Win10) ---
            item_w = 110
            item_h = 110
            spacing = 8  # 必须与 apply_settings 里的 spacing 一致

            # 1. 计算列数：取 (总数) 和 (设置的最大列数) 的较小值
            # 比如总数7个，设置6个 -> 也就是满行6个
            cols = min(count, max_items)

            # 2. 计算行数：向上取整
            # 比如7个，7/6 = 1.16 -> 2行
            rows = math.ceil(count / max_items)

            # 3. 计算宽度 (核心修复点)
            # 宽度 = 列数 * (块宽 + 间距) + 左右边距 + 滚动条预留(即便隐藏)
            # 这里额外 +20 像素作为安全缓冲，防止 Qt 因为差1像素换行
            total_w = cols * (item_w + spacing) + m_left + m_right + 20

            # 4. 计算高度
            # 高度 = 行数 * (块高 + 间距) + 上下边距
            total_h = rows * (item_h + spacing) + m_top + m_bottom

            self.resize(total_w, total_h)

        self.center_window()

    def center_window(self):
        qr = self.frameGeometry()
        screen = QApplication.primaryScreen()
        cp = screen.availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    # --- 显示与切换逻辑 ---

    def show_switcher(self):
        if not self.isVisible():
            self.refresh_windows()

            # 选中第二个（通常是上一个活动窗口），如果是列表末尾则选第0个
            target = 1 if self.list_widget.count() > 1 else 0
            self.list_widget.setCurrentRow(target)

            self.show()
            self.activateWindow()

    def select_next(self):
        if self.list_widget.count() == 0: return
        current = self.list_widget.currentRow()
        next_row = (current + 1) % self.list_widget.count()
        self.list_widget.setCurrentRow(next_row)

    def activate_selected(self):
        if not self.isVisible(): return
        item = self.list_widget.currentItem()
        self.hide()
        if item:
            hwnd = item.data(Qt.ItemDataRole.UserRole)
            self.switch_to_window(hwnd)

    def switch_to_window(self, hwnd):
        """
        核弹级切换窗口：
        1. 模拟按键骗过 Windows 的焦点保护机制。
        2. 处理最小化窗口还原。
        3. 强制夺取前台权限。
        """
        try:
            hwnd = int(hwnd)
            if not win32gui.IsWindow(hwnd):
                return

            # 1. 如果窗口被最小化了，先还原
            # 使用 SW_RESTORE 可以还原最小化的窗口，SW_SHOW 只是显示
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            else:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

            # 2. 【核心黑科技】模拟按下并松开 Alt 键 (VK_MENU = 0x12)
            # 这会欺骗 Windows 认为有物理输入，从而允许当前进程更改前台窗口
            # 0 = KEYEVENTF_EXTENDEDKEY | 0
            # 2 = KEYEVENTF_KEYUP
            user32.keybd_event(0x12, 0, 0, 0)  # Press Alt
            user32.keybd_event(0x12, 0, 2, 0)  # Release Alt

            # 3. 常规切换尝试
            win32gui.SetForegroundWindow(hwnd)
            win32gui.SetFocus(hwnd)

        except Exception as e:
            # print(f"Standard switch failed: {e}, trying brute force...")

            # 4. 如果常规方法失败（通常是 Access Denied），启动暴力模式
            try:
                # 获取当前前台窗口的线程和目标窗口的线程
                foreground_hwnd = win32gui.GetForegroundWindow()
                curr_tid = win32api.GetCurrentThreadId()
                fore_tid, _ = win32process.GetWindowThreadProcessId(foreground_hwnd)
                target_tid, _ = win32process.GetWindowThreadProcessId(hwnd)

                # 将我们的线程“附着”到前台窗口线程上，共享输入队列
                win32process.AttachThreadInput(curr_tid, fore_tid, True)
                if target_tid != fore_tid:
                    win32process.AttachThreadInput(curr_tid, target_tid, True)

                # 再次尝试设置前台
                win32gui.SetForegroundWindow(hwnd)
                win32gui.BringWindowToTop(hwnd)

                # 解除附着
                win32process.AttachThreadInput(curr_tid, fore_tid, False)
                if target_tid != fore_tid:
                    win32process.AttachThreadInput(curr_tid, target_tid, False)

            except:
                # 5. 最后的救命稻草：SwitchToThisWindow
                # 这是个未公开/过时的 API，但在 Win10/11 上对顽固窗口非常有效
                try:
                    user32.SwitchToThisWindow(hwnd, True)
                except:
                    pass


def setup_hook(app_obj):
    # 延迟导入以防止初始化太早
    import keyboard

    def on_alt_tab():
        if not app_obj.isVisible():
            app_obj.sig_show.emit()
        else:
            app_obj.sig_next.emit()

    def on_alt_release(e):
        if app_obj.isVisible():
            app_obj.sig_activate.emit()

    try:
        # suppress=True 会拦截系统的 Alt+Tab
        keyboard.add_hotkey('alt+tab', on_alt_tab, suppress=True)
        keyboard.on_release_key('alt', on_alt_release)
        print("Hook Installed.")
    except Exception as e:
        print(f"Hook Failed: {e}. Ensure you run as Admin.")


if __name__ == "__main__":
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    try:
        ctypes.windll.user32.SystemParametersInfoW(0x2001, 0, ctypes.c_void_p(0), 0x0002 | 0x0001)
    except:
        pass

    switcher = WindowSwitcher()

    # setup_hook(switcher)

    sys.exit(app.exec())