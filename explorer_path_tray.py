# -*- coding: utf-8 -*-
"""
Explorer 路径历史托盘工具

功能:
  - 常驻系统托盘。
  - 后台轮询当前打开的「文件资源管理器」窗口，记录你浏览过的真实文件夹路径。
  - 右键托盘图标 -> 显示最近 10 条路径历史(最新的在最上面)。
  - 点击某条路径 -> 用资源管理器打开它。
  - 弹出菜单后高亮某条路径(鼠标悬停或方向键)并按 Ctrl+C -> 复制该路径。
  - 菜单里「自定义快捷键」-> 弹窗按下任意组合键,即可改掉默认的弹出热键。
  - 历史与配置会保存到本地文件,重启后仍在。

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
import win32clipboard
import pystray
from PIL import Image, ImageDraw


# ---------------------------------------------------------------------------
# 高 DPI 感知
# ---------------------------------------------------------------------------
def _enable_dpi_awareness():
    """声明进程为 DPI 感知,否则高分屏下系统会位图拉伸界面,菜单文字发虚。

    优先用 Per-Monitor v2(Win10 1703+),逐级回退到老接口。必须在创建任何
    窗口之前调用。
    """
    import ctypes
    try:
        # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
        ctypes.windll.user32.SetProcessDpiAwarenessContext(
            ctypes.c_void_p(-4))
        return
    except Exception:
        pass
    try:
        # PROCESS_PER_MONITOR_DPI_AWARE = 2 (Win8.1+)
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()   # 系统级,最老接口
    except Exception:
        pass


_enable_dpi_awareness()


# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------
MAX_HISTORY = 10            # 最多保留多少条路径
POLL_INTERVAL = 1.5         # 轮询间隔(秒)
# 历史保存在脚本所在的项目目录下
HISTORY_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "history.json",
)
# 配置(含自定义快捷键)同样保存在项目目录下
CONFIG_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "config.json",
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
        if not os.path.isdir(path):
            # 路径已不存在 -> 从历史中移除并刷新菜单
            remove_path(path)
            if _icon is not None:
                try:
                    _icon.update_menu()
                except Exception:
                    pass
            return
        # 已有资源管理器窗口时,优先在其中新建标签页打开
        try:
            if _open_in_existing_tab(path):
                return
        except Exception:
            pass
        # 否则(没有已开窗口,或新建标签失败)开新窗口并最大化
        try:
            os.startfile(path)
        except OSError:
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
    yield pystray.MenuItem("自定义快捷键…", _on_set_hotkey)
    yield pystray.MenuItem("退出(&X)", _on_quit)


# ---------------------------------------------------------------------------
# 全局快捷键: Ctrl+Shift+Win+P -> 在鼠标位置弹出菜单
# ---------------------------------------------------------------------------
import ctypes  # noqa: E402
from ctypes import wintypes  # noqa: E402

WM_HOTKEY = 0x0312
WM_APP_REHOTKEY = win32con.WM_APP + 1   # 通知热键线程重新注册热键
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000
HOTKEY_ID = 1
DEFAULT_HOTKEY_MODS = MOD_CONTROL | MOD_SHIFT | MOD_WIN   # 默认 Ctrl+Shift+Win+P
DEFAULT_HOTKEY_VK = 0x50  # 'P'

# 当前生效的热键(可在运行时自定义);_load_config 会用配置文件覆盖
_hotkey_mods = DEFAULT_HOTKEY_MODS | MOD_NOREPEAT
_hotkey_vk = DEFAULT_HOTKEY_VK
_pending_hotkey = None     # 待应用的 (mods, vk),不含 MOD_NOREPEAT
_dialog_open = False       # 自定义快捷键对话框是否已打开

_hotkey_hwnd = None


# ---------------------------------------------------------------------------
# 菜单内按 Ctrl+C 复制高亮项的路径
# ---------------------------------------------------------------------------
# pystray 弹出菜单基于 Win32 TrackPopupMenu,它自带模态消息循环,普通按键
# 不会发到窗口过程。所以这里在菜单显示期间装一个低级键盘钩子来捕获 Ctrl+C,
# 并用 WM_MENUSELECT 实时记录当前高亮的是哪一项。
_menu_highlight_id = 0   # 当前高亮菜单项 id(由 WM_MENUSELECT 维护)
_popup_actions = {}      # 菜单项 id -> (类型, 数据)

WH_KEYBOARD_LL = 13
WM_KEYDOWN = 0x0100
WM_SYSKEYDOWN = 0x0104
VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_MENU = 0x12          # Alt
VK_LWIN = 0x5B
VK_RWIN = 0x5C
VK_C = 0x43

_user32 = ctypes.windll.user32
_HOOKPROC = ctypes.CFUNCTYPE(
    ctypes.c_ssize_t, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)
# 显式声明指针类型,避免 64 位下句柄被截断
_user32.SetWindowsHookExW.restype = ctypes.c_void_p
_user32.SetWindowsHookExW.argtypes = [
    ctypes.c_int, _HOOKPROC, ctypes.c_void_p, wintypes.DWORD]
_user32.CallNextHookEx.restype = ctypes.c_ssize_t
_user32.CallNextHookEx.argtypes = [
    ctypes.c_void_p, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM]
_user32.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]
_user32.UnhookWindowsHookEx.restype = wintypes.BOOL
_user32.RegisterHotKey.argtypes = [
    ctypes.c_void_p, ctypes.c_int, ctypes.c_uint, ctypes.c_uint]
_user32.RegisterHotKey.restype = wintypes.BOOL
_user32.UnregisterHotKey.argtypes = [ctypes.c_void_p, ctypes.c_int]
_user32.UnregisterHotKey.restype = wintypes.BOOL
_user32.GetAsyncKeyState.argtypes = [ctypes.c_int]
_user32.GetAsyncKeyState.restype = ctypes.c_short


class _KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [("vkCode", wintypes.DWORD),
                ("scanCode", wintypes.DWORD),
                ("flags", wintypes.DWORD),
                ("time", wintypes.DWORD),
                ("dwExtraInfo", ctypes.c_void_p)]


_kbd_hook_handle = None
_kbd_hook_proc = None    # 必须保留引用,否则回调被 GC 回收会崩溃


def _copy_to_clipboard(text):
    """把文本写入剪贴板;剪贴板偶尔被占用,所以做几次重试。"""
    for _ in range(5):
        try:
            win32clipboard.OpenClipboard()
        except Exception:
            time.sleep(0.02)
            continue
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
        finally:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass
        return True
    return False


def _copy_path_async(path):
    """在独立线程里复制路径并弹气泡提示(钩子回调里不宜做耗时操作)。"""
    def worker():
        if _copy_to_clipboard(path) and _icon is not None:
            try:
                _icon.notify("已复制路径:\n" + path, "Explorer 路径历史")
            except Exception:
                pass
    threading.Thread(target=worker, daemon=True).start()


def _try_copy_highlighted():
    """复制当前高亮项对应的路径;成功返回 True。"""
    action = _popup_actions.get(_menu_highlight_id)
    if not action or action[0] != "open":
        return False
    _copy_path_async(action[1])
    _user32.EndMenu()    # 复制后关闭菜单,方便用户直接去粘贴
    return True


def _ll_keyboard_proc(nCode, wParam, lParam):
    if nCode == 0 and wParam in (WM_KEYDOWN, WM_SYSKEYDOWN):
        kb = ctypes.cast(lParam, ctypes.POINTER(_KBDLLHOOKSTRUCT)).contents
        if kb.vkCode == VK_C and (_user32.GetAsyncKeyState(VK_CONTROL) & 0x8000):
            if _try_copy_highlighted():
                return 1     # 吞掉这次 Ctrl+C,不再向下传递
    return _user32.CallNextHookEx(None, nCode, wParam, lParam)


def _install_keyboard_hook():
    global _kbd_hook_handle, _kbd_hook_proc
    _kbd_hook_proc = _HOOKPROC(_ll_keyboard_proc)
    _kbd_hook_handle = _user32.SetWindowsHookExW(
        WH_KEYBOARD_LL, _kbd_hook_proc, win32gui.GetModuleHandle(None), 0)


def _uninstall_keyboard_hook():
    global _kbd_hook_handle, _kbd_hook_proc
    if _kbd_hook_handle:
        try:
            _user32.UnhookWindowsHookEx(_kbd_hook_handle)
        except Exception:
            pass
    _kbd_hook_handle = None
    _kbd_hook_proc = None


# ---------------------------------------------------------------------------
# 在已有资源管理器窗口中新建标签页打开路径
# ---------------------------------------------------------------------------
# Windows 11 的资源管理器标签页没有公开 API,只能模拟键盘:把目标窗口切到
# 前台 -> Ctrl+T 新建标签 -> Ctrl+L 定位地址栏 -> 输入路径 -> 回车导航。
# 没有已开窗口或中途失败时,调用方会回退到"开新窗口并最大化"。
VK_RETURN = 0x0D
VK_L = 0x4C
VK_T = 0x54

INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004

# 各步骤之间的等待(秒);太短会因为标签/地址栏还没就绪而丢键
_FG_SETTLE = 0.08      # 切前台后稳定一下
_TAB_NEW = 0.18        # 等新标签建立
_ADDR_FOCUS = 0.06     # 等地址栏获得焦点
_NAV_ENTER = 0.04      # 输入完到回车


class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD),
                ("wScan", wintypes.WORD),
                ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD),
                ("dwExtraInfo", ctypes.c_void_p)]


class _MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", wintypes.LONG),
                ("dy", wintypes.LONG),
                ("mouseData", wintypes.DWORD),
                ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD),
                ("dwExtraInfo", ctypes.c_void_p)]


class _INPUTUNION(ctypes.Union):
    _fields_ = [("ki", _KEYBDINPUT), ("mi", _MOUSEINPUT)]


class _INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUTUNION)]


_user32.SendInput.argtypes = [
    wintypes.UINT, ctypes.POINTER(_INPUT), ctypes.c_int]
_user32.SendInput.restype = wintypes.UINT


def _key_event(vk, scan, flags):
    inp = _INPUT()
    inp.type = INPUT_KEYBOARD
    inp.u.ki = _KEYBDINPUT(vk, scan, flags, 0, None)
    return inp


def _send_inputs(events):
    if not events:
        return
    arr = (_INPUT * len(events))(*events)
    _user32.SendInput(len(events), arr, ctypes.sizeof(_INPUT))


def _send_ctrl_combo(vk):
    """按下 Ctrl + 某键再松开(用于 Ctrl+T / Ctrl+L)。"""
    _send_inputs([
        _key_event(VK_CONTROL, 0, 0),
        _key_event(vk, 0, 0),
        _key_event(vk, 0, KEYEVENTF_KEYUP),
        _key_event(VK_CONTROL, 0, KEYEVENTF_KEYUP),
    ])


def _send_vk(vk):
    _send_inputs([
        _key_event(vk, 0, 0),
        _key_event(vk, 0, KEYEVENTF_KEYUP),
    ])


def _type_unicode(text):
    """逐字符以 Unicode 方式输入,绕开输入法与键盘布局,也不动用户剪贴板。"""
    events = []
    for ch in text:
        code = ord(ch)
        events.append(_key_event(0, code, KEYEVENTF_UNICODE))
        events.append(_key_event(0, code, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP))
    _send_inputs(events)


def _open_explorer_hwnds():
    """返回当前所有资源管理器窗口句柄(去重)。"""
    hwnds = []
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        windows = shell.Windows()
    except Exception:
        return hwnds
    for window in windows:
        try:
            full = (window.FullName or "").lower()
            if not full.endswith("explorer.exe"):
                continue
            hwnd = int(window.HWND)
        except Exception:
            continue
        if hwnd and hwnd not in hwnds and win32gui.IsWindow(hwnd):
            hwnds.append(hwnd)
    return hwnds


def _topmost_hwnd(hwnds):
    """在句柄集合里返回 Z 序最靠前(最近激活)的那个窗口。"""
    if not hwnds:
        return None
    try:
        remaining = set(hwnds)
        hwnd = win32gui.GetTopWindow(0)
        while hwnd:
            if hwnd in remaining:
                return hwnd
            hwnd = win32gui.GetWindow(hwnd, win32con.GW_HWNDNEXT)
    except Exception:
        pass
    return hwnds[0]


def _open_in_existing_tab(path):
    """若已有资源管理器窗口,在其中新建标签页并导航到 path。

    成功返回 True;没有可用窗口或中途失败返回 False(交给调用方回退)。
    """
    hwnds = _open_explorer_hwnds()
    if not hwnds:
        return False
    hwnd = _topmost_hwnd(hwnds)
    if not hwnd:
        return False
    # 最小化的话先恢复,否则切到前台仍不可见,按键也发不进去
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    except Exception:
        pass
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        return False
    time.sleep(_FG_SETTLE)
    # 确认前台确实是目标窗口,否则放弃,避免把按键发到别的程序
    if win32gui.GetForegroundWindow() != hwnd:
        return False
    _send_ctrl_combo(VK_T)         # 新建标签页
    time.sleep(_TAB_NEW)
    if win32gui.GetForegroundWindow() != hwnd:
        return False
    _send_ctrl_combo(VK_L)         # 定位地址栏
    time.sleep(_ADDR_FOCUS)
    _type_unicode(path)            # 输入目标路径
    time.sleep(_NAV_ENTER)
    _send_vk(VK_RETURN)            # 回车导航
    return True


# ---------------------------------------------------------------------------
# 自定义全局快捷键
# ---------------------------------------------------------------------------
# 修饰键的虚拟键码(捕获时用来判断"还没按到主键")
_MODIFIER_VKS = {0x10, 0x11, 0x12, 0x5B, 0x5C,
                 0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5}

# 常见非字母数字键的可读名字
_VK_NAMES = {
    0x08: "Backspace", 0x09: "Tab", 0x0D: "Enter", 0x1B: "Esc",
    0x20: "Space", 0x21: "PageUp", 0x22: "PageDown", 0x23: "End",
    0x24: "Home", 0x25: "Left", 0x26: "Up", 0x27: "Right", 0x28: "Down",
    0x2D: "Insert", 0x2E: "Delete",
}
for _i in range(12):
    _VK_NAMES[0x70 + _i] = "F" + str(_i + 1)
del _i


def _notify(text, title="Explorer 路径历史"):
    if _icon is not None:
        try:
            _icon.notify(text, title)
        except Exception:
            pass


def _vk_name(vk):
    if vk in _VK_NAMES:
        return _VK_NAMES[vk]
    if 0x30 <= vk <= 0x5A:        # 0-9 与 A-Z 的 VK 即其 ASCII
        return chr(vk)
    return "VK_{:02X}".format(vk)


def _format_hotkey(mods, vk):
    """把 (修饰位, 主键) 格式化成 'Ctrl + Shift + P' 这样的可读文字。"""
    parts = []
    if mods & MOD_CONTROL:
        parts.append("Ctrl")
    if mods & MOD_SHIFT:
        parts.append("Shift")
    if mods & MOD_ALT:
        parts.append("Alt")
    if mods & MOD_WIN:
        parts.append("Win")
    if vk:
        parts.append(_vk_name(vk))
    return " + ".join(parts) if parts else "(未设置)"


def _current_modifiers():
    """读取此刻按下的修饰键,返回 RegisterHotKey 用的 MOD_* 组合。"""
    g = _user32.GetAsyncKeyState
    mods = 0
    if g(VK_CONTROL) & 0x8000:
        mods |= MOD_CONTROL
    if g(VK_SHIFT) & 0x8000:
        mods |= MOD_SHIFT
    if g(VK_MENU) & 0x8000:
        mods |= MOD_ALT
    if (g(VK_LWIN) & 0x8000) or (g(VK_RWIN) & 0x8000):
        mods |= MOD_WIN
    return mods


def _load_config():
    global _hotkey_mods, _hotkey_vk
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        hk = data.get("hotkey") or {}
        mods = int(hk["mods"])
        vk = int(hk["vk"])
        _hotkey_mods = (mods & ~MOD_NOREPEAT) | MOD_NOREPEAT
        _hotkey_vk = vk
    except (OSError, ValueError, TypeError, KeyError):
        _hotkey_mods = DEFAULT_HOTKEY_MODS | MOD_NOREPEAT
        _hotkey_vk = DEFAULT_HOTKEY_VK


def _save_config():
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"hotkey": {"mods": _hotkey_mods & ~MOD_NOREPEAT,
                                  "vk": _hotkey_vk}},
                      f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def apply_hotkey(mods, vk):
    """记录待应用的热键,并通知热键线程去重新注册(注册必须在该线程做)。"""
    global _pending_hotkey
    _pending_hotkey = (mods, vk)
    if _hotkey_hwnd:
        win32gui.PostMessage(_hotkey_hwnd, WM_APP_REHOTKEY, 0, 0)


def _reregister_hotkey():
    """在热键线程内执行:用待应用的组合重新注册全局热键,失败则回滚。"""
    global _hotkey_mods, _hotkey_vk
    if _pending_hotkey is None or not _hotkey_hwnd:
        return
    mods, vk = _pending_hotkey
    new_mods = mods | MOD_NOREPEAT
    _user32.UnregisterHotKey(_hotkey_hwnd, HOTKEY_ID)
    if _user32.RegisterHotKey(_hotkey_hwnd, HOTKEY_ID, new_mods, vk):
        _hotkey_mods, _hotkey_vk = new_mods, vk
        _save_config()
        _notify("快捷键已设为 " + _format_hotkey(mods, vk))
    else:
        # 新组合注册失败(多半被占用),回滚到之前能用的热键
        _user32.RegisterHotKey(_hotkey_hwnd, HOTKEY_ID, _hotkey_mods, _hotkey_vk)
        _notify("设置失败:" + _format_hotkey(mods, vk) + " 可能已被占用")


def _capture_hotkey_dialog():
    """弹出小窗口捕获用户按下的组合键,确认后应用为新的全局快捷键。"""
    global _dialog_open
    if _dialog_open:
        return
    _dialog_open = True
    try:
        import tkinter as tk
    except Exception:
        _notify("无法打开设置窗口(缺少 tkinter)")
        _dialog_open = False
        return

    pending = {"mods": _hotkey_mods & ~MOD_NOREPEAT, "vk": _hotkey_vk}

    root = tk.Tk()
    root.title("自定义快捷键")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    tk.Label(root,
             text="请按下想要的组合键\n(需至少包含一个修饰键 Ctrl / Shift / Alt / Win)",
             justify="center", padx=24, pady=10).pack()

    combo_var = tk.StringVar(value=_format_hotkey(pending["mods"], pending["vk"]))
    tk.Label(root, textvariable=combo_var,
             font=("Segoe UI", 15, "bold"), fg="#0a7", pady=4).pack()

    hint_var = tk.StringVar(value="")
    tk.Label(root, textvariable=hint_var, fg="#c00").pack()

    def on_keypress(event):
        vk = event.keycode
        if vk in _MODIFIER_VKS:      # 只按了修饰键,等主键
            return "break"
        mods = _current_modifiers()
        pending["mods"] = mods
        pending["vk"] = vk
        combo_var.set(_format_hotkey(mods, vk))
        hint_var.set("" if mods else "缺少修饰键,可能无法注册")
        return "break"

    root.bind("<KeyPress>", on_keypress)

    btns = tk.Frame(root)
    btns.pack(pady=12)

    def save():
        if pending["vk"] and pending["mods"]:
            apply_hotkey(pending["mods"], pending["vk"])
            root.destroy()
        else:
            hint_var.set("请至少包含一个修饰键再按主键")

    def cancel():
        root.destroy()

    tk.Button(btns, text="保存", width=8, command=save).pack(side="left", padx=8)
    tk.Button(btns, text="取消", width=8, command=cancel).pack(side="left", padx=8)

    # 居中显示
    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry("+{}+{}".format((sw - w) // 2, (sh - h) // 2))

    root.lift()
    root.focus_force()
    try:
        root.mainloop()
    finally:
        try:
            root.destroy()
        except Exception:
            pass
        _dialog_open = False


def _on_set_hotkey(icon=None, item=None):
    # tkinter 在独立线程跑自己的 mainloop,避免阻塞托盘
    threading.Thread(target=_capture_hotkey_dialog, daemon=True).start()


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
    global _popup_actions, _menu_highlight_id
    hmenu = win32gui.CreatePopupMenu()
    _popup_actions = {}  # 菜单项 id -> (类型, 数据)
    _menu_highlight_id = 0
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
            _popup_actions[next_id] = ("open", path)
            next_id += 1
    win32gui.AppendMenu(hmenu, win32con.MF_SEPARATOR, 0, "")
    win32gui.AppendMenu(hmenu, win32con.MF_STRING, next_id, "自定义快捷键…")
    _popup_actions[next_id] = ("set_hotkey", None)
    next_id += 1
    win32gui.AppendMenu(hmenu, win32con.MF_STRING, next_id, "退出(&X)")
    _popup_actions[next_id] = ("quit", None)

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
    # 菜单显示期间装上键盘钩子,这样高亮某项时按 Ctrl+C 就能复制其路径
    _install_keyboard_hook()
    try:
        cmd = win32gui.TrackPopupMenu(
            hmenu, align | win32con.TPM_RETURNCMD, x, y, 0, hwnd, None)
    finally:
        _uninstall_keyboard_hook()
    win32gui.PostMessage(hwnd, win32con.WM_NULL, 0, 0)
    win32gui.DestroyMenu(hmenu)

    action = _popup_actions.get(cmd)
    if action:
        kind, data = action
        if kind == "open":
            open_path(data)
        elif kind == "set_hotkey":
            _on_set_hotkey()
        elif kind == "quit":
            _quit_app()


def _hotkey_wndproc(hwnd, msg, wparam, lparam):
    global _menu_highlight_id
    if msg == WM_HOTKEY and wparam == HOTKEY_ID:
        _show_popup_menu(hwnd)
        return 0
    if msg == WM_APP_REHOTKEY:
        _reregister_hotkey()
        return 0
    if msg == win32con.WM_MENUSELECT:
        # wParam: 低字=菜单项 id(或子菜单索引), 高字=菜单标志
        item_id = wparam & 0xFFFF
        flags = (wparam >> 16) & 0xFFFF
        if (flags == 0xFFFF and lparam == 0) or (flags & win32con.MF_POPUP):
            _menu_highlight_id = 0   # 菜单已关闭或高亮的是子菜单
        else:
            _menu_highlight_id = item_id
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

        if not _user32.RegisterHotKey(
                hwnd, HOTKEY_ID, _hotkey_mods, _hotkey_vk):
            # 注册失败(多半是热键已被占用),提示一下;不退出线程,
            # 这样用户仍可通过菜单把它改成别的可用组合。
            _notify("快捷键 "
                    + _format_hotkey(_hotkey_mods & ~MOD_NOREPEAT, _hotkey_vk)
                    + " 注册失败(可能已被占用),可在菜单里改成别的组合")

        try:
            win32gui.PumpMessages()
        finally:
            _user32.UnregisterHotKey(hwnd, HOTKEY_ID)
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
    _load_config()

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
