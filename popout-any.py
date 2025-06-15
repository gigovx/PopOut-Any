import sys
import os
import win32gui, win32con, win32api
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QListWidget, QPushButton,
    QSpinBox, QLabel, QCheckBox, QVBoxLayout, QHBoxLayout,
    QSystemTrayIcon, QMenu
)
from PyQt6.QtGui import QCursor, QIcon, QAction, QPixmap
from PyQt6.QtCore import QTimer, Qt

def resource_path(rel_path: str) -> str:
    """
    Get absolute path to resource, works for dev and for PyInstaller bundle.
    """
    base = getattr(sys, "_MEIPASS", os.path.dirname(__file__))
    return os.path.join(base, rel_path)

# The eight edge‐segments, clockwise
SEGMENTS = [
    "Top-Left", "Top-Right",
    "Right-Top", "Right-Bottom",
    "Bottom-Right", "Bottom-Left",
    "Left-Bottom", "Left-Top"
]
# Map segment → slide direction index (0=Left,1=Right,2=Top,3=Bottom)
DIRECTION_MAP = {
    "Top-Left":     2, "Top-Right":    2,
    "Right-Top":    1, "Right-Bottom": 1,
    "Bottom-Right": 3, "Bottom-Left":  3,
    "Left-Bottom":  0, "Left-Top":     0
}

def enum_visible_windows():
    wins = []
    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        # skip child windows
        if win32gui.GetParent(hwnd) != 0:
            return
        # skip some shell helpers
        cls = win32gui.GetClassName(hwnd)
        if cls in ("Progman", "Shell_TrayWnd", "Windows.UI.Core.CoreWindow"):
            return
        # require caption & sysmenu
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        if not (style & win32con.WS_CAPTION and style & win32con.WS_SYSMENU):
            return
        # skip tiny
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        if (r - l) < 100 or (b - t) < 40:
            return
        title = win32gui.GetWindowText(hwnd)
        if title:
            wins.append((hwnd, title))
    win32gui.EnumWindows(cb, None)
    return wins

class Slider:
    """Animate a window with optional easing/fade/activation."""
    def __init__(self, hwnd, start, end, steps, interval,
                 use_ease=False, use_fade=False,
                 into_view=False, activate_on_show=False, cb=None):
        self.hwnd = hwnd
        self.sx, self.sy = start
        self.ex, self.ey = end
        self.steps = max(1, steps)
        self.interval = interval
        self.use_ease = use_ease
        self.use_fade = use_fade
        self.into_view = into_view
        self.activate_on_show = activate_on_show
        self.cb = cb
        self.i = 0
        if use_fade:
            ex = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            win32gui.SetWindowLong(
                hwnd, win32con.GWL_EXSTYLE,
                ex | win32con.WS_EX_LAYERED
            )
        self.timer = QTimer()
        self.timer.setInterval(self.interval)
        self.timer.timeout.connect(self._step)

    def start(self):
        self.i = 0
        self.timer.start()

    def _step(self):
        self.i += 1
        t = self.i / self.steps
        if self.use_ease:
            t = t*t*(3 - 2*t)
        nx = int(self.sx + (self.ex - self.sx)*t)
        ny = int(self.sy + (self.ey - self.sy)*t)
        flags = win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE
        win32gui.SetWindowPos(self.hwnd, None, nx, ny, 0, 0, flags)
        if self.use_fade:
            alpha = int(255*t) if self.into_view else int(255*(1-t))
            win32gui.SetLayeredWindowAttributes(self.hwnd, 0, alpha, win32con.LWA_ALPHA)
        if self.i >= self.steps:
            self.timer.stop()
            if self.activate_on_show and self.into_view:
                win32gui.SetWindowPos(
                    self.hwnd, None,
                    self.ex, self.ey, 0, 0,
                    win32con.SWP_NOSIZE | win32con.SWP_NOZORDER
                )
                try:
                    win32gui.SetForegroundWindow(self.hwnd)
                except Exception:
                    pass
            if self.cb:
                self.cb()

class SlideAnyWindowApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PopOut-Any")

        # set both main-window and tray icon
        icon = QIcon(resource_path("icon.ico"))
        self.setWindowIcon(icon)

        # core state
        self.window_cfg      = {seg: [] for seg in SEGMENTS}
        self.segment_active  = None
        self.interacting     = False
        self.animators       = []
        self.segment_buttons = {}
        self.active_segment  = None

        # build UI
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        # 1) Window picker
        layout.addWidget(QLabel("1) Select window"))
        self.win_list = QListWidget()
        self.win_list.setMinimumHeight(200)
        self.win_list.itemSelectionChanged.connect(self._on_sel)
        layout.addWidget(self.win_list)
        btn_refresh = QPushButton("Refresh Windows")
        btn_refresh.clicked.connect(self._populate_windows)
        layout.addWidget(btn_refresh)

        # 2) Clickable edge-selector image
        layout.addWidget(QLabel("2) Select Edge Segment"))
        self.segment_label = QLabel()
        pix = QPixmap(resource_path("monitor_base.png")).scaled(
            353, 252,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation
        )
        self.segment_label.setPixmap(pix)
        self.segment_label.setFixedSize(353, 252)
        layout.addWidget(self.segment_label, alignment=Qt.AlignmentFlag.AlignCenter)

        # carve out 8 transparent buttons
        w, h      = 353, 252
        thickness = int(h * 0.10)
        half_h    = h // 2

        def add_segment(name, x, y, w_, h_):
            btn = QPushButton(self.segment_label)
            btn.setGeometry(x, y, w_, h_)
            btn.setObjectName(name)
            btn.setStyleSheet("""
                QPushButton { background: rgba(0,0,0,0); border: none; }
                QPushButton:hover { background: rgba(0,174,239,0.2); }
            """)
            btn.clicked.connect(lambda _, s=name: self._on_segment_clicked(s))
            self.segment_buttons[name] = btn

        # top
        add_segment("Top-Left",    0,           0,        w//3,          thickness)
        add_segment("Top-Right",   2*(w//3),    0,        w - 2*(w//3), thickness)
        # bottom
        add_segment("Bottom-Left", 0,       h-thickness, w//3,          thickness)
        add_segment("Bottom-Right",2*(w//3), h-thickness, w - 2*(w//3), thickness)
        # left
        add_segment("Left-Top",    0,           thickness, thickness, half_h - thickness)
        add_segment("Left-Bottom", 0,           half_h,    thickness, half_h)
        # right
        add_segment("Right-Top",   w-thickness, thickness, thickness, half_h - thickness)
        add_segment("Right-Bottom",w-thickness, half_h,    thickness, half_h)

        # 3) Speed & Steps
        layout.addWidget(QLabel("3) Speed & Steps"))
        row = QHBoxLayout()
        row.addWidget(QLabel("Interval ms"))
        self.speed = QSpinBox(); self.speed.setRange(5,500); self.speed.setValue(15)
        row.addWidget(self.speed)
        row.addStretch(1)
        row.addWidget(QLabel("Steps"))
        self.steps = QSpinBox(); self.steps.setRange(5,200); self.steps.setValue(30)
        row.addWidget(self.steps)
        layout.addLayout(row)

        # 4) Animation options
        layout.addWidget(QLabel("4) Animation"))
        self.chk_ease = QCheckBox("Use Easing")
        self.chk_fade = QCheckBox("Fade Window")
        layout.addWidget(self.chk_ease)
        layout.addWidget(self.chk_fade)

        # 5) Assign
        self.assign_btn = QPushButton("Assign to Segment")
        self.assign_btn.clicked.connect(self._assign)
        layout.addWidget(self.assign_btn)

        # 6) Current assignments
        layout.addWidget(QLabel("6) Current assignments"))
        self.assign_list = QListWidget()
        self.assign_list.itemSelectionChanged.connect(self._on_assign_sel)
        layout.addWidget(self.assign_list)
        self.remove_btn = QPushButton("Remove Selected")
        self.remove_btn.clicked.connect(self._remove_assignment)
        layout.addWidget(self.remove_btn)

        # 7) Enable edge trigger
        self.enable_btn = QPushButton("Enable Edge Trigger")
        self.enable_btn.setCheckable(True)
        self.enable_btn.toggled.connect(self._on_enable)
        layout.addWidget(self.enable_btn)

        # disable until valid
        self.assign_btn.setEnabled(False)
        self.remove_btn.setEnabled(False)
        self.enable_btn.setEnabled(False)

        # timers
        self.edge_timer = QTimer(self);   self.edge_timer.setInterval(100)
        self.edge_timer.timeout.connect(self._check_cursor_edge)
        self.focus_timer = QTimer(self);  self.focus_timer.setInterval(200)
        self.focus_timer.timeout.connect(self._check_focus)

        # system tray
        self.tray_icon = QSystemTrayIcon(icon, parent=self)
        tray_menu = QMenu()
        tray_menu.addAction(QAction("Restore", self, triggered=self._restore_from_tray))
        tray_menu.addAction(QAction("Exit",    self, triggered=QApplication.quit))
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self._on_tray_activated)

        # initial load
        self._populate_windows()

    # — Segment click handler ———————————————————————————————————————
    def _on_segment_clicked(self, name):
        self.active_segment = name
        for seg, btn in self.segment_buttons.items():
            active = (seg == name)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: rgba({0 if active else 0},
                                           {174 if active else 0},
                                           {239 if active else 0},
                                           {0.3 if active else 0});
                    border: none;
                }}
                QPushButton:hover {{
                    background-color: rgba(0,174,239,0.2);
                }}
            """)
        self._on_sel()

    # — Enable/disable controls —————————————————————————————————————
    def _on_sel(self):
        has_win = bool(self.win_list.selectedItems())
        has_seg = (self.active_segment is not None)
        self.assign_btn.setEnabled(has_win and has_seg)
        self.enable_btn.setEnabled(any(self.window_cfg.values()))

    # — Window list —————————————————————————————————————————————————
    def _populate_windows(self):
        self.win_list.clear()
        for hwnd, title in enum_visible_windows():
            self.win_list.addItem(f"{hwnd} | {title}")
        self._on_sel()

    # — Assign ——————————————————————————————————————————————————————
    def _assign(self):
        item = self.win_list.currentItem()
        if not item or not self.active_segment:
            return
        hwnd = int(item.text().split("|",1)[0].strip())
        seg = self.active_segment
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        cfg = {
            "hwnd":     hwnd,
            "orig":     (l, t, r, b),
            "steps":    self.steps.value(),
            "interval": self.speed.value(),
            "ease":     self.chk_ease.isChecked(),
            "fade":     self.chk_fade.isChecked()
        }
        if not any(c["hwnd"] == hwnd for c in self.window_cfg[seg]):
            self.window_cfg[seg].append(cfg)
        self._refresh_assignments()

    def _refresh_assignments(self):
        self.assign_list.clear()
        for seg, lst in self.window_cfg.items():
            for c in lst:
                title = win32gui.GetWindowText(c["hwnd"])
                self.assign_list.addItem(f"{seg} | {title} | {c['hwnd']}")
        self._on_sel()

    def _on_assign_sel(self):
        can = bool(self.assign_list.selectedItems()) and not self.enable_btn.isChecked()
        self.remove_btn.setEnabled(can)

    def _remove_assignment(self):
        sel = self.assign_list.currentItem()
        if not sel:
            return
        seg, _, hwnd_str = sel.text().rsplit("|", 2)
        seg = seg.strip()
        hwnd = int(hwnd_str.strip())
        self.window_cfg[seg] = [c for c in self.window_cfg[seg] if c["hwnd"] != hwnd]
        self._refresh_assignments()

    # — Edge trigger ON/OFF ————————————————————————————————————
    def _on_enable(self, on: bool):
        if on:
            for seg, lst in self.window_cfg.items():
                for cfg in lst:
                    self._hide_taskbar(cfg["hwnd"])
                    self._slide_cfg(cfg, seg, into_view=False)
            self.segment_active = None
            self.interacting = False
            self.edge_timer.start()
            self.focus_timer.start()
            self.remove_btn.setEnabled(False)
        else:
            self.edge_timer.stop()
            self.focus_timer.stop()
            def restore(cfg):
                self._show_taskbar(cfg["hwnd"])
            for seg, lst in self.window_cfg.items():
                for cfg in lst:
                    self._slide_cfg(cfg, seg, into_view=True,
                                    cb=lambda cfg=cfg: restore(cfg))
            self.interacting = False
            self._on_assign_sel()

    # — Core window methods ————————————————————————————————————————
    def _hide_taskbar(self, hwnd):
        ex = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        ex = (ex & ~win32con.WS_EX_APPWINDOW) | win32con.WS_EX_TOOLWINDOW
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex)
        win32gui.SetWindowPos(
            hwnd, None, 0, 0, 0, 0,
            win32con.SWP_NOMOVE|
            win32con.SWP_NOSIZE|
            win32con.SWP_NOZORDER|
            win32con.SWP_FRAMECHANGED
        )

    def _show_taskbar(self, hwnd):
        ex = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        ex = (ex | win32con.WS_EX_APPWINDOW) & ~win32con.WS_EX_TOOLWINDOW
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex)
        win32gui.SetWindowPos(
            hwnd, None, 0, 0, 0, 0,
            win32con.SWP_NOMOVE|
            win32con.SWP_NOSIZE|
            win32con.SWP_NOZORDER|
            win32con.SWP_FRAMECHANGED
        )

    def _compute_positions(self, hwnd, dir_idx, orig):
        mon = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)
        mi = win32api.GetMonitorInfo(mon)["Monitor"]
        lm, tm, rm, bm = mi
        l, t, r, b = orig
        w, h = r - l, b - t
        pad = 2
        left_x, right_x = lm - w - pad, rm + pad
        top_y, bottom_y = tm - h - pad, bm + pad
        if dir_idx == 0:   return ((left_x, t), (l, t))
        if dir_idx == 1:   return ((right_x, t), (l, t))
        if dir_idx == 2:   return ((l, top_y), (l, t))
        return ((l, bottom_y), (l, t))

    def _slide_cfg(self, cfg, seg, into_view, cb=None):
        idx = DIRECTION_MAP[seg]
        start, end = self._compute_positions(cfg["hwnd"], idx, cfg["orig"])
        if not into_view:
            start, end = end, start

        def done():
            if into_view:
                self.interacting = True
                # 🔹 Ensure window moves **above full-screen** windows!
                win32gui.SetWindowPos(
                    cfg["hwnd"], win32con.HWND_TOPMOST,
                    end[0], end[1], 0, 0,
                    win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
                )

                # 🔹 Force focus ONLY if necessary
                try:
                    win32gui.SetForegroundWindow(cfg["hwnd"])
                except Exception:
                    pass
            
            if cb:
                cb()

        anim = Slider(
            cfg["hwnd"], start, end,
            cfg["steps"], cfg["interval"],
            use_ease=cfg["ease"], use_fade=cfg["fade"],
            into_view=into_view,
            activate_on_show=into_view,
            cb=done
        )
        self.animators.append(anim)
        anim.start()

    def _get_seg_hit(self, x, y):
        geom = QApplication.primaryScreen().geometry()
        sw, sh = geom.width(), geom.height()
        m = 3
        midx, midy = sw//2, sh//2
        if y <= m:       return 0 if x <= midx else 1
        if y >= sh - m:  return 5 if x <= midx else 4
        if x <= m:       return 7 if y <= midy else 6
        if x >= sw - m:  return 2 if y <= midy else 3
        return None

    def _check_cursor_edge(self):
        if not self.enable_btn.isChecked() or self.interacting or self.active_segment is None:
            return
        pos = QCursor.pos()
        hit = self._get_seg_hit(pos.x(), pos.y())
        want = SEGMENTS.index(self.active_segment)
        if hit == want and self.segment_active is None:
            self.segment_active = want
            for cfg in self.window_cfg[self.active_segment]:
                self._slide_cfg(cfg, self.active_segment, into_view=True)

    def _check_focus(self):
        if not self.enable_btn.isChecked() or not self.interacting:
            return
        fg = win32gui.GetForegroundWindow()
        seg = SEGMENTS[self.segment_active]
        if not any(c["hwnd"] == fg for c in self.window_cfg[seg]):
            for c in self.window_cfg[seg]:
                self._slide_cfg(c, seg, into_view=False)
            self.segment_active = None
            self.interacting = False

    def changeEvent(self, ev):
        super().changeEvent(ev)
        if ev.type() == ev.Type.WindowStateChange and self.isMinimized():
            QTimer.singleShot(0, self.hide)
            self.tray_icon.show()

    def _on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self._restore_from_tray()

    def _restore_from_tray(self):
        self.showNormal()
        self.activateWindow()
        self.tray_icon.hide()

    def closeEvent(self, ev):
        if self.enable_btn.isChecked():
            self.enable_btn.setChecked(False)
        for seg in SEGMENTS:
            for cfg in self.window_cfg[seg]:
                l, t, r, b = cfg["orig"]
                win32gui.SetWindowPos(
                    cfg["hwnd"], None, l, t, 0, 0,
                    win32con.SWP_NOSIZE | win32con.SWP_NOZORDER
                )
                self._show_taskbar(cfg["hwnd"])
        ev.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = SlideAnyWindowApp()
    win.resize(700, 650)
    win.show()
    sys.exit(app.exec())