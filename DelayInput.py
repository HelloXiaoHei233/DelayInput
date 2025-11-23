import sys
import time
import threading
import random
import os

import pyautogui
import keyboard

from PyQt6.QtCore import (
    Qt,
    QTimer,
    pyqtSignal,
    QObject,
    QEvent,
)
from PyQt6.QtGui import (
    QFont,
    QKeySequence,
    QColor,
    QPalette,
)
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QHBoxLayout,
    QVBoxLayout,
    QGroupBox,
    QTextEdit,
    QLabel,
    QSlider,
    QSpinBox,
    QAbstractSpinBox,
    QCheckBox,
    QPushButton,
    QProgressBar,
    QToolButton,
    QMessageBox,
    QLineEdit,
)


# ================= 自定义：可拖入文件的文本框 =================

class DroppableTextEdit(QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        md = event.mimeData()
        if md.hasUrls() or md.hasText():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        md = event.mimeData()
        # 优先处理文件
        if md.hasUrls():
            urls = md.urls()
            if urls:
                path = urls[0].toLocalFile()
                if path:
                    self._load_file(path)
                    event.acceptProposedAction()
                    return
        # 其他情况下按普通文本处理
        super().dropEvent(event)

    def _load_file(self, path: str):
        try:
            size_bytes = os.path.getsize(path)
            size_str = self._format_size(size_bytes)

            # 大小判断：超过 1000 字节弹窗确认
            if size_bytes > 1000:
                reply = QMessageBox.question(
                    self,
                    "文件较大",
                    f"该文件超过了1000字（文件大小：{size_str}），你确定要继续吗？",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if reply != QMessageBox.StandardButton.Yes:
                    return

            with open(path, "rb") as f:
                data = f.read()

            # 尝试用 utf-8 解码，无法解码的部分丢弃
            try:
                text = data.decode("utf-8")
            except UnicodeDecodeError:
                text = data.decode("utf-8", errors="ignore")

            self.setPlainText(text)

        except Exception as e:
            QMessageBox.critical(self, "读取失败", f"无法读取文件：\n{e}")

    @staticmethod
    def _format_size(num_bytes: int) -> str:
        size = float(num_bytes)
        for unit in ["B", "KB", "MB", "GB", "TB"]:
            if size < 1024:
                if unit == "B":
                    return f"{int(size)}{unit}"
                else:
                    return f"{size:.2f}{unit}"
            size /= 1024.0
        return f"{size:.2f}PB"


# ================= 自定义：快捷键编辑框 =================

class HotkeyEdit(QLineEdit):
    """
    捕获组合键：
    - 聚焦后开始捕获
    - 失焦或 Enter 提交
    - Esc 取消还原
    """
    sequenceCommitted = pyqtSignal(str, bool)  # (sequence_str, has_main_key)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self._capturing = False
        self._current_sequence = ""
        self._has_main_key = False
        self._original_text = ""

    def _update_focus_style(self):
        # 聚焦时高亮边框，失焦时还原默认
        if self._capturing:
            self.setStyleSheet(
                "border-radius: 6px;"
                "border: 1px solid #3B82F6;"
                "padding: 4px 8px;"
                "background: #020617;"
            )
        else:
            self.setStyleSheet("")

    def focusInEvent(self, event):
        super().focusInEvent(event)
        self._capturing = True
        self._current_sequence = ""
        self._has_main_key = False
        self._original_text = self.text()
        self.selectAll()
        self._update_focus_style()

    def focusOutEvent(self, event):
        super().focusOutEvent(event)
        was_capturing = self._capturing
        self._capturing = False
        self._update_focus_style()
        if was_capturing:
            self.sequenceCommitted.emit(self._current_sequence, self._has_main_key)

    def keyPressEvent(self, event):
        if not self._capturing:
            super().keyPressEvent(event)
            return

        key = event.key()

        # Enter：结束并提交
        if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.clearFocus()
            return

        # Esc：取消捕获，还原
        if key == Qt.Key.Key_Escape:
            self._current_sequence = ""
            self._has_main_key = False
            self.setText(self._original_text)
            self.clearFocus()
            return

        modifiers = []
        mods = event.modifiers()
        if mods & Qt.KeyboardModifier.ControlModifier:
            modifiers.append("Ctrl")
        if mods & Qt.KeyboardModifier.ShiftModifier:
            modifiers.append("Shift")
        if mods & Qt.KeyboardModifier.AltModifier:
            modifiers.append("Alt")
        if mods & Qt.KeyboardModifier.MetaModifier:
            modifiers.append("Win")

        main_key = ""
        if key not in (
            Qt.Key.Key_Control,
            Qt.Key.Key_Shift,
            Qt.Key.Key_Alt,
            Qt.Key.Key_Meta,
        ):
            seq = QKeySequence(key)
            main_key = seq.toString()
            if len(main_key) == 1:
                main_key = main_key.upper()
            self._has_main_key = True

        parts = modifiers.copy()
        if main_key:
            parts.append(main_key)

        display = "+".join(parts)
        self._current_sequence = display
        self.setText(display)

    def setOccupied(self, occupied: bool):
        pal = self.palette()
        if occupied:
            pal.setColor(QPalette.ColorRole.Text, QColor("#EF4444"))
        else:
            pal.setColor(QPalette.ColorRole.Text, QColor("#D0D4DA"))
        self.setPalette(pal)


# ================= 工具：打字线程 =================

class TypingWorker(QObject):
    progress_changed = pyqtSignal(int)  # 0~100
    finished = pyqtSignal()
    stopped = pyqtSignal()
    error = pyqtSignal(str)
    focus_paused = pyqtSignal()  # 焦点变化导致的自动暂停

    def __init__(self, text, base_delay_ms, use_random, rand_min_ms, rand_max_ms, target_window_title=None):
        super().__init__()
        self.text = text
        self.base_delay_ms = base_delay_ms
        self.use_random = use_random
        self.rand_min_ms = rand_min_ms
        self.rand_max_ms = rand_max_ms
        if self.rand_max_ms < self.rand_min_ms:
            self.rand_min_ms, self.rand_max_ms = self.rand_max_ms, self.rand_min_ms

        self._stop_flag = False
        self._pause_flag = False

        self.target_window_title = target_window_title

    def stop(self):
        self._stop_flag = True

    def pause(self):
        self._pause_flag = True

    def resume(self):
        self._pause_flag = False

    def set_target_window(self, title):
        """更新目标窗口标题，解决切换窗口后无法恢复的问题"""
        self.target_window_title = title

    @staticmethod
    def _is_fast_char(ch: str) -> bool:
        if ch in ("\n", "\t"):
            return False
        return ord(ch) >= 32

    def _safe_write(self, text: str):
        try:
            keyboard.write(text, delay=0)
        except Exception:
            pyautogui.write(text, interval=0)

    def _fast_type_chunk(self, chunk: str):
        if chunk:
            self._safe_write(chunk)

    def _check_focus(self):
        if not self.target_window_title or self._stop_flag or self._pause_flag:
            return

        try:
            current_title = pyautogui.getActiveWindowTitle()
        except Exception:
            return

        # 无焦点或切换到其它窗口则自动暂停
        if not current_title or current_title != self.target_window_title:
            self._pause_flag = True
            self.focus_paused.emit()

    def run(self):
        try:
            length = len(self.text)
            if length == 0:
                self.finished.emit()
                return

            i = 0
            fast_mode = (self.base_delay_ms == 0 and not self.use_random)
            max_chunk = 40

            while i < length:
                if self._stop_flag:
                    self.stopped.emit()
                    return

                self._check_focus()

                while self._pause_flag and not self._stop_flag:
                    time.sleep(0.05)

                if self._stop_flag:
                    self.stopped.emit()
                    return

                if fast_mode and self._is_fast_char(self.text[i]):
                    chunk_chars = []
                    while (
                        i < length
                        and len(chunk_chars) < max_chunk
                        and self._is_fast_char(self.text[i])
                    ):
                        chunk_chars.append(self.text[i])
                        i += 1
                    self._fast_type_chunk("".join(chunk_chars))
                else:
                    ch = self.text[i]
                    self._type_char(ch)
                    i += 1

                percent = int(i / length * 100)
                self.progress_changed.emit(percent)

                if not fast_mode:
                    delay_ms = self.base_delay_ms
                    if self.use_random:
                        delay_ms += random.randint(self.rand_min_ms, self.rand_max_ms)
                    if delay_ms > 0:
                        time.sleep(delay_ms / 1000.0)

            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

    def _type_char(self, ch: str):
        # 换行
        if ch == "\n":
            try:
                keyboard.write("\n", delay=0)
            except Exception:
                pyautogui.press("enter")
            return

        # 制表符
        if ch == "\t":
            try:
                keyboard.write("\t", delay=0)
            except Exception:
                pyautogui.press("tab")
            return

        code = ord(ch)
        # 简单处理几个常见控制字符
        if code < 32:
            if code == 8:
                pyautogui.press("backspace")
            elif code == 27:
                pyautogui.press("esc")
            return

        # 普通字符 / 中文 / emoji
        self._safe_write(ch)


# ================= UI：主窗口 =================

class MainWindow(QMainWindow):
    STATE_IDLE = 0
    STATE_COUNTDOWN = 1
    STATE_TYPING = 2
    STATE_PAUSED = 3

    def __init__(self):
        super().__init__()
        self.setWindowTitle("延迟输入工具")
        self.resize(980, 620)

        self.state = self.STATE_IDLE
        self.current_countdown_ms = 0
        self.current_resume_ms = 0

        self.typing_thread = None
        self.typing_worker = None

        self.hotkey_str = "ctrl+shift+t"
        self.hotkey_handle = None
        self.hotkey_occupied = False

        self._build_ui()
        self._apply_md3_style()

        QApplication.instance().installEventFilter(self)

        self._register_hotkey()
        self._set_idle("等待操作", 0)

    # ---------- 公共小工具 ----------

    def _reset_timers(self):
        self.countdown_timer.stop()
        self.resume_timer.stop()

    def _set_idle(self, text="等待操作", progress=None, is_error=False):
        self._reset_timers()
        self.state = self.STATE_IDLE
        self.btn_pause.setText("暂停")
        self.btn_start.setEnabled(True)
        self._update_status(text, progress, is_error)

    # ---------- 事件过滤：快捷键输入框失焦 ----------

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.MouseButtonPress:
            if getattr(self, "hotkey_edit", None) is not None:
                if self.hotkey_edit.hasFocus() and obj is not self.hotkey_edit:
                    self.hotkey_edit.clearFocus()
        return super().eventFilter(obj, event)

    # ---------- UI 搭建 ----------

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(8)

        # 顶部标题 + 置顶
        title_bar = QHBoxLayout()
        title_label = QLabel("延迟输入工具")
        title_label.setStyleSheet("font-size: 18px; font-weight: 600;")
        title_bar.addWidget(title_label)
        title_bar.addStretch()

        self.pin_button = QToolButton()
        self.pin_button.setText("置顶")
        self.pin_button.setToolTip("置顶 / 取消置顶")
        self.pin_button.setCheckable(True)
        self.pin_button.clicked.connect(self._toggle_on_top)
        title_bar.addWidget(self.pin_button)

        root_layout.addLayout(title_bar)

        # 中间主区域
        main_layout = QHBoxLayout()
        main_layout.setSpacing(12)
        root_layout.addLayout(main_layout, 1)

        # 左：文本输入区
        left_group = QGroupBox()
        left_layout = QVBoxLayout(left_group)
        self.text_edit = DroppableTextEdit()
        self.text_edit.setPlaceholderText("粘贴文字或拖入待复制文件......")
        font = QFont("Consolas" if "Consolas" in QFont().families() else self.font().family(), 11)
        self.text_edit.setFont(font)
        left_layout.addWidget(self.text_edit)

        main_layout.addWidget(left_group, 2)

        # 右：控制区
        right_panel = QVBoxLayout()
        right_panel.setSpacing(10)
        main_layout.addLayout(right_panel, 3)

        # 打字设置
        settings_group = QGroupBox()
        settings_layout = QVBoxLayout(settings_group)

        # 速度
        speed_layout = QHBoxLayout()
        speed_label = QLabel("打字速度：")
        self.speed_slider = QSlider(Qt.Orientation.Horizontal)
        self.speed_slider.setRange(0, 2000)
        self.speed_slider.setValue(0)
        self.speed_slider.valueChanged.connect(self._on_speed_slider_changed)

        self.speed_spin = QSpinBox()
        self.speed_spin.setRange(0, 2000)
        self.speed_spin.setValue(0)
        self.speed_spin.setSuffix(" ms")
        self.speed_spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.speed_spin.valueChanged.connect(self._on_speed_spin_changed)

        speed_layout.addWidget(speed_label)
        speed_layout.addWidget(self.speed_slider, 1)
        speed_layout.addWidget(self.speed_spin)
        settings_layout.addLayout(speed_layout)

        # 随机延迟 + 开始/继续延迟
        rand_layout = QHBoxLayout()
        self.random_checkbox = QCheckBox("随机延迟")
        rand_layout.addWidget(self.random_checkbox)

        rand_layout.addWidget(QLabel("最小："))
        self.rand_min_spin = QSpinBox()
        self.rand_min_spin.setRange(0, 2000)
        self.rand_min_spin.setValue(5)
        self.rand_min_spin.setSuffix(" ms")
        self.rand_min_spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.rand_min_spin.setFixedWidth(70)
        rand_layout.addWidget(self.rand_min_spin)

        rand_layout.addWidget(QLabel("最大："))
        self.rand_max_spin = QSpinBox()
        self.rand_max_spin.setRange(0, 2000)
        self.rand_max_spin.setValue(10)
        self.rand_max_spin.setSuffix(" ms")
        self.rand_max_spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.rand_max_spin.setFixedWidth(70)
        rand_layout.addWidget(self.rand_max_spin)

        rand_layout.addWidget(QLabel("开始/继续延迟："))
        self.start_delay_spin = QSpinBox()
        self.start_delay_spin.setRange(0, 60000)
        self.start_delay_spin.setValue(3000)
        self.start_delay_spin.setSuffix(" ms")
        self.start_delay_spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        self.start_delay_spin.setFixedWidth(90)
        rand_layout.addWidget(self.start_delay_spin)

        settings_layout.addLayout(rand_layout)
        right_panel.addWidget(settings_group)

        # 控制按钮
        ctrl_group = QGroupBox("控制")
        ctrl_layout = QVBoxLayout(ctrl_group)

        button_row = QHBoxLayout()
        self.btn_start = QPushButton("开始")
        self.btn_pause = QPushButton("暂停")
        self.btn_stop = QPushButton("中止")

        self.btn_start.clicked.connect(self._on_start_clicked)
        self.btn_pause.clicked.connect(self._on_pause_clicked)
        self.btn_stop.clicked.connect(self._on_stop_clicked)

        button_row.addWidget(self.btn_start)
        button_row.addWidget(self.btn_pause)
        button_row.addWidget(self.btn_stop)
        ctrl_layout.addLayout(button_row)
        right_panel.addWidget(ctrl_group)

        # 快捷键
        shortcut_group = QGroupBox("快捷键")
        shortcut_layout = QHBoxLayout(shortcut_group)

        shortcut_label = QLabel("启动打字：")
        shortcut_layout.addWidget(shortcut_label)

        self.hotkey_edit = HotkeyEdit()
        self.hotkey_edit.setFixedWidth(220)
        self.hotkey_edit.sequenceCommitted.connect(self._on_hotkey_committed)
        shortcut_layout.addWidget(self.hotkey_edit)

        right_panel.addWidget(shortcut_group)

        # 状态 & 进度
        status_group = QGroupBox("状态")
        status_layout = QVBoxLayout(status_group)

        self.status_label = QLabel("当前状态：等待操作")
        status_layout.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        status_layout.addWidget(self.progress_bar)

        right_panel.addWidget(status_group)
        right_panel.addStretch(1)

        # 计时器
        self.countdown_timer = QTimer(self)
        self.countdown_timer.timeout.connect(self._on_countdown_tick)

        self.resume_timer = QTimer(self)
        self.resume_timer.timeout.connect(self._on_resume_tick)

    # ---------- 样式 ----------

    def _apply_md3_style(self):
        app = QApplication.instance()
        if app is not None:
            app.setStyle("Fusion")

        self.setStyleSheet("""
            QMainWindow {
                background-color: #101316;
            }
            QGroupBox {
                border: 1px solid #2A2E33;
                border-radius: 12px;
                margin-top: 18px;
                padding: 10px;
                background-color: #181C20;
                font-weight: 500;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 6px;
                color: #D0D4DA;
            }
            QTextEdit {
                border-radius: 10px;
                border: 1px solid #30363D;
                padding: 8px;
                background: #0F1115;
                color: #E3E7EB;
                selection-background-color: #3B82F6;
            }
            QLabel {
                color: #D0D4DA;
            }
            QSlider::groove:horizontal {
                border: 1px solid #30363D;
                height: 6px;
                border-radius: 3px;
                background: #1F2933;
            }
            QSlider::handle:horizontal {
                background: #3B82F6;
                border-radius: 8px;
                width: 18px;
                margin: -5px 0;
            }
            QSlider::sub-page:horizontal {
                background: #2563EB;
                border-radius: 3px;
            }
            QSpinBox {
                border-radius: 6px;
                border: 1px solid #30363D;
                padding: 2px 6px;
                background: #0F1115;
                color: #E3E7EB;
            }
            QCheckBox {
                spacing: 6px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 6px;
                border: 1px solid #3B82F6;
                background: transparent;
            }
            QCheckBox::indicator:checked {
                background: #3B82F6;
            }
            QPushButton {
                border-radius: 999px;
                border: 1px solid #3B82F6;
                padding: 8px 14px;
                background: #1F2937;
                color: #E3E7EB;
            }
            QPushButton:hover {
                background: #2B3543;
            }
            QPushButton:pressed {
                background: #111827;
            }
            QPushButton:disabled {
                background: #1F2937;
                color: #6B7280;
                border-color: #4B5563;
            }
            QProgressBar {
                border-radius: 8px;
                border: 1px solid #30363D;
                background: #0F1115;
                text-align: center;
                color: #E3E7EB;
            }
            QProgressBar::chunk {
                background-color: #22C55E;
                border-radius: 8px;
            }
            QToolButton {
                border-radius: 999px;
                padding: 4px 8px;
                background: #1F2933;
                border: 1px solid #30363D;
            }
            QToolButton:checked {
                background: #3B82F6;
                border-color: #3B82F6;
                color: #FFFFFF;
            }
            QLineEdit {
                border-radius: 6px;
                border: 1px solid #30363D;
                padding: 4px 8px;
                background: #0F1115;
            }
        """)

        self._set_hotkey_display(self.hotkey_occupied)

    # ---------- 窗口置顶 ----------

    def _toggle_on_top(self, checked):
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, checked)
        self.show()

    # ---------- 打字速度控件联动 ----------

    def _on_speed_slider_changed(self, value):
        if self.speed_spin.value() != value:
            self.speed_spin.setValue(value)

    def _on_speed_spin_changed(self, value):
        if self.speed_slider.value() != value:
            self.speed_slider.setValue(value)

    # ---------- 热键显示/注册 ----------

    @staticmethod
    def _to_readable_hotkey(hotkey_str: str) -> str:
        parts = []
        for raw in hotkey_str.split("+"):
            raw = raw.strip()
            if not raw:
                continue
            low = raw.lower()
            if low in ("ctrl", "control"):
                parts.append("Ctrl")
            elif low == "shift":
                parts.append("Shift")
            elif low in ("alt", "menu"):
                parts.append("Alt")
            elif low in ("win", "windows", "meta", "super"):
                parts.append("Win")
            else:
                parts.append(raw.upper() if len(raw) == 1 else raw.capitalize())
        return "+".join(parts)

    def _display_hotkey_text(self, hotkey_str: str, occupied: bool) -> str:
        base = self._to_readable_hotkey(hotkey_str)
        return f"[被占用]{base}" if occupied else base

    def _set_hotkey_display(self, occupied: bool):
        self.hotkey_occupied = occupied
        text = self._display_hotkey_text(self.hotkey_str, occupied)
        self.hotkey_edit.setOccupied(occupied)
        self.hotkey_edit.setText(text)

    @staticmethod
    def _canonicalize_sequence(seq_str: str) -> str:
        parts = []
        for p in seq_str.split("+"):
            p = p.strip()
            if not p:
                continue
            low = p.lower()
            if low in ("ctrl", "control"):
                parts.append("ctrl")
            elif low == "shift":
                parts.append("shift")
            elif low in ("alt", "menu"):
                parts.append("alt")
            elif low in ("win", "windows", "meta", "super"):
                parts.append("windows")
            else:
                parts.append(low)
        return "+".join(parts)

    def _register_hotkey(self):
        if self.hotkey_handle is not None:
            try:
                keyboard.remove_hotkey(self.hotkey_handle)
            except Exception:
                pass
            self.hotkey_handle = None

        try:
            self.hotkey_handle = keyboard.add_hotkey(self.hotkey_str, self._on_hotkey_trigger)
            self._set_hotkey_display(occupied=False)
        except Exception:
            self._set_hotkey_display(occupied=True)

    def _on_hotkey_committed(self, seq_str: str, has_main_key: bool):
        seq_str = seq_str.strip()

        # 没有按任何键 / 只有修饰键
        if not seq_str or not has_main_key:
            self._set_hotkey_display(self.hotkey_occupied)
            return

        parts = [p.strip() for p in seq_str.split("+") if p.strip()]
        # 至少两个键
        if len(parts) < 2:
            self._set_hotkey_display(self.hotkey_occupied)
            return

        # 以 Ctrl/Shift/Alt 开头
        first = parts[0].lower()
        if first not in ("ctrl", "control", "shift", "alt"):
            self._set_hotkey_display(self.hotkey_occupied)
            return

        # 最后一个必须是字母/数字/单字符符号或 F1~F24
        main_display = parts[-1]
        main_low = main_display.lower()
        allowed = False
        if len(main_display) == 1:
            allowed = True
        elif main_low.startswith("f") and main_low[1:].isdigit():
            allowed = True
        if not allowed:
            self._set_hotkey_display(self.hotkey_occupied)
            return

        new_hotkey = self._canonicalize_sequence(seq_str)

        # 没变化
        if new_hotkey == self.hotkey_str:
            self._set_hotkey_display(self.hotkey_occupied)
            return

        self.hotkey_str = new_hotkey
        self._register_hotkey()

    def _on_hotkey_trigger(self):
        QTimer.singleShot(0, self._on_hotkey_trigger_gui)

    def _on_hotkey_trigger_gui(self):
        if getattr(self, "hotkey_edit", None) is not None:
            he = self.hotkey_edit
            if he.hasFocus() or getattr(he, "_capturing", False):
                return
        self._start_typing_immediately_from_hotkey()

    # ---------- 状态更新 ----------

    def _update_status(self, text, progress=None, is_error=False):
        if progress is None:
            progress = self.progress_bar.value()
        self.status_label.setText(f"当前状态：{text}")
        self.progress_bar.setValue(progress)
        self.status_label.setStyleSheet("color: #EF4444;" if is_error else "color: #D0D4DA;")

    # ---------- 控制按钮逻辑 ----------

    def _on_start_clicked(self):
        if self.state in (self.STATE_TYPING, self.STATE_COUNTDOWN):
            self._cancel_typing()
        elif self.state == self.STATE_PAUSED:
            self._cancel_typing()
            self._start_typing(skip_countdown=False)
        else:
            self._start_typing(skip_countdown=False)

    def _on_pause_clicked(self):
        if self.state == self.STATE_TYPING and self.typing_worker:
            self.resume_timer.stop()
            self.typing_worker.pause()
            self.state = self.STATE_PAUSED
            self.btn_pause.setText("继续")
            self._update_status("已暂停", self.progress_bar.value())
        elif self.state == self.STATE_PAUSED and self.typing_worker:
            if self.resume_timer.isActive():
                self.resume_timer.stop()
                self._update_status("已暂停", self.progress_bar.value())
            else:
                delay_ms = self.start_delay_spin.value()
                if delay_ms <= 0:
                    self._resume_typing(from_hotkey=False)
                else:
                    self.current_resume_ms = delay_ms
                    self.resume_timer.start(100)
                    sec = self.current_resume_ms / 1000.0
                    self._update_status(f"{sec:.1f} 秒后继续输入...", self.progress_bar.value())

    def _on_stop_clicked(self):
        self._cancel_typing()

    # ---------- 暂停后恢复 ----------

    def _on_resume_tick(self):
        step = 100
        self.current_resume_ms -= step
        if self.current_resume_ms <= 0:
            self.resume_timer.stop()
            self._resume_typing(from_hotkey=False)
        else:
            sec = self.current_resume_ms / 1000.0
            self._update_status(f"{sec:.1f} 秒后继续输入...", self.progress_bar.value())

    def _resume_typing(self, from_hotkey: bool = False):
        if not (self.typing_worker and self.state == self.STATE_PAUSED):
            return

        try:
            current_win = pyautogui.getActiveWindowTitle()
        except Exception:
            current_win = None

        if not current_win:
            self._update_status("无法获取输入焦点(无窗口) - 恢复失败", self.progress_bar.value(), is_error=True)
            self.resume_timer.stop()
            return

        if current_win == self.windowTitle():
            self._update_status("输入焦点不能是本工具 - 继续暂停中", self.progress_bar.value(), is_error=True)
            self.resume_timer.stop()
            return

        # 更新目标窗口为当前窗口
        self.typing_worker.set_target_window(current_win)

        self.typing_worker.resume()
        self.state = self.STATE_TYPING
        self.btn_pause.setText("暂停")
        tip = "输入中（快捷键继续）" if from_hotkey else "输入中"
        self._update_status(tip, self.progress_bar.value())

    # ---------- 开始 / 倒计时 / 取消 ----------

    def _start_typing(self, skip_countdown=False):
        if not self.text_edit.toPlainText():
            self._update_status("没有要输入的文本", self.progress_bar.value(), is_error=True)
            return

        self._update_status("准备开始输入", self.progress_bar.value(), is_error=False)

        if skip_countdown:
            self._begin_typing()
            return

        delay_ms = self.start_delay_spin.value()
        if delay_ms <= 0:
            self._begin_typing()
        else:
            self.state = self.STATE_COUNTDOWN
            self.current_countdown_ms = delay_ms
            sec = self.current_countdown_ms / 1000.0
            self._update_status(
                f"倒计时中：{sec:.1f} 秒后开始输入",
                self.progress_bar.value(),
            )
            self.countdown_timer.start(100)
            self.btn_start.setEnabled(False)

    def _on_countdown_tick(self):
        if self.state != self.STATE_COUNTDOWN:
            self.countdown_timer.stop()
            self.btn_start.setEnabled(True)
            return

        step = 100
        self.current_countdown_ms -= step
        if self.current_countdown_ms <= 0:
            self.countdown_timer.stop()
            self._begin_typing()
        else:
            sec = self.current_countdown_ms / 1000.0
            self._update_status(
                f"倒计时中：{sec:.1f} 秒后开始输入",
                self.progress_bar.value(),
            )

    def _begin_typing(self):
        try:
            target_window_title = pyautogui.getActiveWindowTitle()
        except Exception:
            target_window_title = None

        if not target_window_title:
            self._set_idle("输入焦点不正确（无窗口焦点）- 启动失败", 0, is_error=True)
            return

        if self.isActiveWindow():
            self._set_idle("输入焦点不正确 - 请先切换到目标窗口", 0, is_error=True)
            return

        self.state = self.STATE_TYPING
        self.btn_pause.setText("暂停")
        self.resume_timer.stop()
        self._update_status("输入中", self.progress_bar.value())
        self.btn_start.setEnabled(False)

        text = self.text_edit.toPlainText()
        base_delay_ms = self.speed_spin.value()
        use_random = self.random_checkbox.isChecked()
        rand_min = self.rand_min_spin.value()
        rand_max = self.rand_max_spin.value()

        self.typing_worker = TypingWorker(
            text,
            base_delay_ms,
            use_random,
            rand_min,
            rand_max,
            target_window_title=target_window_title,
        )

        self.typing_worker.progress_changed.connect(self._on_typing_progress)
        self.typing_worker.finished.connect(self._on_typing_finished)
        self.typing_worker.stopped.connect(self._on_typing_stopped)
        self.typing_worker.error.connect(self._on_typing_error)
        self.typing_worker.focus_paused.connect(self._on_focus_paused)

        self.typing_thread = threading.Thread(target=self.typing_worker.run, daemon=True)
        self.typing_thread.start()

    def _cancel_typing(self):
        if self.state == self.STATE_IDLE:
            return

        if self.state == self.STATE_COUNTDOWN:
            self._set_idle("已取消", self.progress_bar.value())
        elif self.state in (self.STATE_TYPING, self.STATE_PAUSED):
            if self.typing_worker:
                self.typing_worker.stop()
            self._set_idle("已中止", self.progress_bar.value())

    def _start_typing_immediately_from_hotkey(self):
        if self.state == self.STATE_PAUSED and self.typing_worker:
            self._resume_typing(from_hotkey=True)
        elif self.state in (self.STATE_TYPING, self.STATE_COUNTDOWN):
            self._cancel_typing()
        else:
            self._start_typing(skip_countdown=True)

    # ---------- 打字线程回调 ----------

    def _on_typing_progress(self, percent):
        self._update_status(f"输入中（{percent}%）", percent)

    def _on_typing_finished(self):
        self._set_idle("输入完成", 100)

    def _on_typing_stopped(self):
        self._set_idle("已中止", self.progress_bar.value())

    def _on_typing_error(self, msg):
        self._set_idle(f"[错误] {msg}", self.progress_bar.value(), is_error=True)
        QMessageBox.critical(self, "错误", f"输入过程中出现错误：{msg}")

    def _on_focus_paused(self):
        if self.state not in (self.STATE_TYPING, self.STATE_PAUSED):
            return

        self.resume_timer.stop()
        self.state = self.STATE_PAUSED
        self.btn_pause.setText("继续")

        try:
            current_win = pyautogui.getActiveWindowTitle()
        except Exception:
            current_win = None

        if not current_win:
            self._update_status("输入焦点不正确（无窗口焦点）- 已暂停", self.progress_bar.value(), is_error=True)
        elif current_win == self.windowTitle():
            self._update_status("输入焦点不能是本工具 - 继续暂停中", self.progress_bar.value(), is_error=True)
        else:
            self._update_status("窗口焦点变化，已临时暂停", self.progress_bar.value())

    # ---------- 关闭 ----------

    def closeEvent(self, event):
        try:
            if self.hotkey_handle is not None:
                keyboard.remove_hotkey(self.hotkey_handle)
        except Exception:
            pass
        event.accept()


# ================= main =================

def main():
    pyautogui.FAILSAFE = False

    app = QApplication(sys.argv)
    app.setApplicationName("延迟输入工具")

    win = MainWindow()
    win.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
