# -*- coding: utf-8 -*-
"""
Explorer 路径历史托盘工具

功能:
  - 常驻系统托盘。
  - 后台轮询当前打开的「文件资源管理器」窗口，记录你浏览过的真实文件夹路径。
  - 右键托盘图标 -> 显示最近 10 条路径历史(最新的在最上面)。
  - 点击某条路径 -> 用资源管理器打开它。
  - 历史会保存到本地文件,重启后仍在。

依赖: pywin32, pystray, Pillow
运行: pythonw explorer_path_tray.py   (用 pythonw 不弹黑窗口)
"""

import os
import sys
import json
import time
import threading

import pythoncom
import win32com.client
import win32gui
import win32con
import pystray
from PIL import Image, ImageDraw

# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------
MAX_HISTORY = 10            # 最多保留多少条路径
POLL_INTERVAL = 1.5         # 轮询间隔(秒)
HISTORY_FILE = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
    "ExplorerPathTray",
    "history.json",
)

# ---------------------------------------------------------------------------
# 历史记录(线程安全)
# ---------------------------------------------------------------------------
_lock = threading.Lock()
_history = []          # 最近在前,元素为真实路径字符串
_icon = None           # pystray.Icon,稍后赋值


def _load_history():
    global _history
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            _history = [p for p in data if isinstance(p, str)][:MAX_HISTORY]
    except (OSError, ValueError):
        _history = []


def _save_history():
    try:
        os.makedirs(os.path.dirname(HISTORY_FILE), exist_ok=True)
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(_history, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def add_path(path):
    """把一个路径放到历史最前面;返回 True 表示历史发生了变化。"""
    if not path:
        return False
    changed = False
    with _lock:
        # 按不区分大小写去重,但保留首次出现的原始写法
        existing = next((p for p in _history if p.lower() == path.lower()), None)
        if existing is None:
            _history.insert(0, path)
            changed = True
        elif _history[0] != existing:
            _history.remove(existing)
            _history.insert(0, existing)
            changed = True
        del _history[MAX_HISTORY:]
    if changed:
        _save_history()
    return changed


def remove_path(path):
    with _lock:
        before = len(_history)
        _history[:] = [p for p in _history if p.lower() != path.lower()]
        changed = len(_history) != before
    if changed:
        _save_history()
    return changed


def history_snapshot():
    with _lock:
        return list(_history)


# ---------------------------------------------------------------------------
# 读取当前打开的资源管理器窗口路径
# ---------------------------------------------------------------------------
def get_open_explorer_paths():
    """返回当前所有「文件资源管理器」窗口正在浏览的真实文件夹路径。"""
    paths = []
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        windows = shell.Windows()
    except Exception:
        return paths

    for window in windows:
        try:
            # 只要资源管理器窗口,排除 IE 等其它 ShellWindows 成员
            full = (window.FullName or "").lower()
            if not full.endswith("explorer.exe"):
                continue
            # Self.Path 给出真实文件系统路径;对“此电脑/控制面板”等会是 ::{GUID}
            path = window.Document.Folder.Self.Path
        except Exception:
            continue
        if path and os.path.isdir(path):
            paths.append(os.path.normpath(path))
    return paths


def poll_loop():
    """后台线程:周期性记录正在浏览的路径。"""
    pythoncom.CoInitialize()
    try:
        stop = _stop_event
        while not stop.is_set():
            changed = False
            for p in get_open_explorer_paths():
                if add_path(p):
                    changed = True
            if changed and _icon is not None:
                try:
                    _icon.update_menu()
                except Exception:
                    pass
            stop.wait(POLL_INTERVAL)
    finally:
        pythoncom.CoUninitialize()


_stop_event = threading.Event()


# ---------------------------------------------------------------------------
# 托盘菜单
# ---------------------------------------------------------------------------
def _maximize_explorer_window(target_path, timeout=3.0):
    """找到正在显示 target_path 的资源管理器窗口并将其最大化、置前。"""
    target = os.path.normcase(os.path.normpath(target_path))
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            windows = shell.Windows()
        except Exception:
            windows = []
        for window in windows:
            try:
                full = (window.FullName or "").lower()
                if not full.endswith("explorer.exe"):
                    continue
                path = window.Document.Folder.Self.Path
                if not path:
                    continue
                if os.path.normcase(os.path.normpath(path)) != target:
                    continue
                hwnd = int(window.HWND)
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                try:
                    win32gui.SetForegroundWindow(hwnd)
                except Exception:
                    pass
                return True
            except Exception:
                continue
        time.sleep(0.1)
    return False


def _open_and_maximize(path):
    """在独立线程中打开路径并最大化窗口(独立线程便于隔离 COM)。"""
    pythoncom.CoInitialize()
    try:
        try:
            os.startfile(path)
        except OSError:
            # 路径已不存在 -> 从历史中移除并刷新菜单
            remove_path(path)
            if _icon is not None:
                try:
                    _icon.update_menu()
                except Exception:
                    pass
            return
        _maximize_explorer_window(path)
    finally:
        pythoncom.CoUninitialize()


def open_path(path):
    threading.Thread(target=_open_and_maximize, args=(path,), daemon=True).start()


def _make_open(path):
    def handler(icon, item):
        open_path(path)
    return handler


def _quit_app():
    _stop_event.set()
    if _icon is not None:
        _icon.stop()


def _on_quit(icon, item):
    _quit_app()


def menu_items():
    """动态生成菜单项,每次右键弹出时都会重新求值。"""
    snapshot = history_snapshot()
    if not snapshot:
        yield pystray.MenuItem("(暂无浏览记录)", None, enabled=False)
    else:
        for path in snapshot:
            # 菜单文字里把单个 & 转义,避免被当成助记符吞掉
            label = path.replace("&", "&&")
            yield pystray.MenuItem(label, _make_open(path))
    yield pystray.Menu.SEPARATOR
    yield pystray.MenuItem("退出(&X)", _on_quit)


# ---------------------------------------------------------------------------
# 全局快捷键: Ctrl+Shift+Win+P -> 在鼠标位置弹出菜单
# ---------------------------------------------------------------------------
import ctypes  # noqa: E402

WM_HOTKEY = 0x0312
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000
HOTKEY_ID = 1
HOTKEY_MODS = MOD_CONTROL | MOD_SHIFT | MOD_WIN | MOD_NOREPEAT
HOTKEY_VK = 0x50  # 'P'

_hotkey_hwnd = None


class _GUID(ctypes.Structure):
    _fields_ = [("Data1", ctypes.c_ulong),
                ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort),
                ("Data4", ctypes.c_ubyte * 8)]


class _NOTIFYICONIDENTIFIER(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_ulong),
                ("hWnd", ctypes.c_void_p),
                ("uID", ctypes.c_uint),
                ("guidItem", _GUID)]


class _RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


def _get_tray_icon_rect():
    """返回 pystray 托盘图标在屏幕上的矩形(失败返回 None)。"""
    if _icon is None:
        return None
    hwnd = getattr(_icon, "_hwnd", None)
    if not hwnd:
        return None
    ident = _NOTIFYICONIDENTIFIER()
    ident.cbSize = ctypes.sizeof(_NOTIFYICONIDENTIFIER)
    ident.hWnd = int(hwnd)
    # pystray 注册图标时 uID 实际为 0(其源码误把字段名写成 hID,被 ctypes 忽略)
    ident.uID = 0
    rect = _RECT()
    try:
        hr = ctypes.windll.shell32.Shell_NotifyIconGetRect(
            ctypes.byref(ident), ctypes.byref(rect))
    except Exception:
        return None
    if hr != 0:  # S_OK == 0
        return None
    return rect


def _show_popup_menu(hwnd):
    """在鼠标位置弹出与右键托盘一致的菜单。"""
    hmenu = win32gui.CreatePopupMenu()
    actions = {}  # 菜单项 id -> (类型, 数据)
    next_id = 1
    snapshot = history_snapshot()
    if not snapshot:
        win32gui.AppendMenu(
            hmenu, win32con.MF_STRING | win32con.MF_GRAYED, next_id, "(暂无浏览记录)")
        next_id += 1
    else:
        for path in snapshot:
            # '&' 在菜单文字里是助记符,需转义成 '&&' 才会原样显示
            win32gui.AppendMenu(
                hmenu, win32con.MF_STRING, next_id, path.replace("&", "&&"))
            actions[next_id] = ("open", path)
            next_id += 1
    win32gui.AppendMenu(hmenu, win32con.MF_SEPARATOR, 0, "")
    win32gui.AppendMenu(hmenu, win32con.MF_STRING, next_id, "退出(&X)")
    actions[next_id] = ("quit", None)

    # 菜单位置:锚定到托盘图标右上角,并用与右键一致的右下对齐,
    # 让菜单贴着图标向左上展开;取不到图标位置时回退到鼠标处。
    rect = _get_tray_icon_rect()
    if rect is not None:
        x, y = rect.right, rect.top
    else:
        x, y = win32gui.GetCursorPos()
    align = win32con.TPM_RIGHTALIGN | win32con.TPM_BOTTOMALIGN
    # 必须先把(隐藏)窗口设为前台,菜单才能在点击别处时正常消失
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    cmd = win32gui.TrackPopupMenu(
        hmenu, align | win32con.TPM_RETURNCMD, x, y, 0, hwnd, None)
    win32gui.PostMessage(hwnd, win32con.WM_NULL, 0, 0)
    win32gui.DestroyMenu(hmenu)

    action = actions.get(cmd)
    if action:
        kind, data = action
        if kind == "open":
            open_path(data)
        elif kind == "quit":
            _quit_app()


def _hotkey_wndproc(hwnd, msg, wparam, lparam):
    if msg == WM_HOTKEY and wparam == HOTKEY_ID:
        _show_popup_menu(hwnd)
        return 0
    if msg == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0
    return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)


def hotkey_loop():
    """后台线程:创建隐藏窗口、注册全局热键并跑消息循环。"""
    global _hotkey_hwnd
    pythoncom.CoInitialize()
    try:
        hinst = win32gui.GetModuleHandle(None)
        wc = win32gui.WNDCLASS()
        wc.lpszClassName = "ExplorerPathTrayHotkey"
        wc.lpfnWndProc = _hotkey_wndproc
        wc.hInstance = hinst
        class_atom = win32gui.RegisterClass(wc)
        hwnd = win32gui.CreateWindow(
            class_atom, "ExplorerPathTrayHotkey", 0, 0, 0, 0, 0, 0, 0, hinst, None)
        _hotkey_hwnd = hwnd

        if not ctypes.windll.user32.RegisterHotKey(
                hwnd, HOTKEY_ID, HOTKEY_MODS, HOTKEY_VK):
            # 注册失败(多半是热键已被占用),通过气泡提示一下
            if _icon is not None:
                try:
                    _icon.notify("快捷键 Ctrl+Shift+Win+P 注册失败(可能已被占用)",
                                 "Explorer 路径历史")
                except Exception:
                    pass
            return

        try:
            win32gui.PumpMessages()
        finally:
            ctypes.windll.user32.UnregisterHotKey(hwnd, HOTKEY_ID)
    finally:
        pythoncom.CoUninitialize()


# ---------------------------------------------------------------------------
# 托盘图标
# ---------------------------------------------------------------------------
ICON_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "2048x2048.png")


def make_icon_image():
    """加载项目目录下的 png 作为托盘图标;失败时回退到自绘文件夹图标。"""
    try:
        img = Image.open(ICON_FILE)
        img.load()
        # 托盘图标不需要这么大,缩放到 64x64
        return img.convert("RGBA").resize((64, 64), Image.LANCZOS)
    except Exception:
        pass
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    body = (8, 22, 56, 52)
    tab = (8, 14, 30, 24)
    d.rounded_rectangle(tab, radius=3, fill=(232, 184, 64, 255))
    d.rounded_rectangle(body, radius=4, fill=(255, 205, 84, 255))
    return img


def main():
    global _icon
    _load_history()

    def setup(icon):
        icon.visible = True
        threading.Thread(target=poll_loop, daemon=True).start()
        threading.Thread(target=hotkey_loop, daemon=True).start()

    _icon = pystray.Icon(
        "ExplorerPathTray",
        icon=make_icon_image(),
        title="Explorer 路径历史",
        menu=pystray.Menu(menu_items),
    )
    _icon.run(setup=setup)


if __name__ == "__main__":
    main()
