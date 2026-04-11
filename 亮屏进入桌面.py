"""
亮屏进入桌面 - 唤醒屏幕并模拟 Enter 键进入桌面
✅ 增强版：桌面就绪检测 + 多次 Enter 键尝试 + 强制激活桌面焦点（解决计划任务点击失效）
✅ 超时优化：总执行时间 < 15 秒，避免外部调用超时（30秒）
✅ 兼容计划任务：通过 SetForegroundWindow 激活桌面，确保 SendInput 生效
"""

import ctypes
from ctypes import wintypes
import time
import sys

# ---------- 可选：输出日志到文件（便于计划任务调试）----------
# LOG_FILE = r"C:\temp\unlock_log.txt"
# sys.stdout = open(LOG_FILE, "w", encoding="utf-8")

# ---------- 常量定义 ----------
INPUT_MOUSE = 0
INPUT_KEYBOARD = 1
MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_ABSOLUTE = 0x8000
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
KEYEVENTF_KEYUP = 0x0002
VK_RETURN = 0x0D
VK_MENU = 0x12          # Alt 键，用于激活输入焦点

# ---------- 结构体定义 ----------
class MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG),
                ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT)]
    _anonymous_ = ("_input",)
    _fields_ = [("type", wintypes.DWORD), ("_input", _INPUT)]

# ---------- 基础输入函数 ----------
def send_mouse_move(dx, dy):
    """模拟鼠标相对移动"""
    mi = MOUSEINPUT(dx=dx, dy=dy, dwFlags=MOUSEEVENTF_MOVE)
    inp = INPUT(type=INPUT_MOUSE, mi=mi)
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))

def send_key(vk_code, press=True):
    """模拟按键（按下或弹起）"""
    flags = 0 if press else KEYEVENTF_KEYUP
    ki = KEYBDINPUT(wVk=vk_code, dwFlags=flags)
    inp = INPUT(type=INPUT_KEYBOARD, ki=ki)
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))

def send_mouse_click(x=None, y=None, button="left"):
    """模拟鼠标点击，支持移动到绝对坐标"""
    if x is not None and y is not None:
        screen_width = ctypes.windll.user32.GetSystemMetrics(0)
        screen_height = ctypes.windll.user32.GetSystemMetrics(1)
        dx = int(x * 65535 / screen_width)
        dy = int(y * 65535 / screen_height)
        mi = MOUSEINPUT(dx=dx, dy=dy, dwFlags=MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE)
        inp = INPUT(type=INPUT_MOUSE, mi=mi)
        ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
        time.sleep(0.05)

    if button == "left":
        flags_down = MOUSEEVENTF_LEFTDOWN
        flags_up = MOUSEEVENTF_LEFTUP
    else:
        flags_down = MOUSEEVENTF_RIGHTDOWN
        flags_up = MOUSEEVENTF_RIGHTUP

    mi_down = MOUSEINPUT(dx=0, dy=0, dwFlags=flags_down)
    inp_down = INPUT(type=INPUT_MOUSE, mi=mi_down)
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp_down), ctypes.sizeof(inp_down))
    time.sleep(0.05)

    mi_up = MOUSEINPUT(dx=0, dy=0, dwFlags=flags_up)
    inp_up = INPUT(type=INPUT_MOUSE, mi=mi_up)
    ctypes.windll.user32.SendInput(1, ctypes.byref(inp_up), ctypes.sizeof(inp_up))

# ---------- 焦点激活函数（解决计划任务环境输入失效）----------
def activate_desktop():
    """激活桌面窗口（Progman 或 Shell_TrayWnd），强制获得输入焦点"""
    # 查找桌面窗口句柄（Progman 是桌面壁纸宿主窗口，Shell_TrayWnd 是任务栏）
    hwnd = ctypes.windll.user32.FindWindowW("Progman", None)
    if not hwnd:
        hwnd = ctypes.windll.user32.FindWindowW("Shell_TrayWnd", None)
    if hwnd:
        # 显示窗口（如果最小化）
        ctypes.windll.user32.ShowWindow(hwnd, 5)      # SW_SHOW
        # 设置为前台窗口
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(0.3)
        # 发送一个无害的 Alt 键，帮助激活输入状态
        send_key(VK_MENU, press=True)
        time.sleep(0.05)
        send_key(VK_MENU, press=False)
        time.sleep(0.4)
        return True
    return False

# ---------- 状态检测函数 ----------
def is_desktop_ready():
    """检测桌面是否就绪（Explorer 进程存在）"""
    try:
        handle = ctypes.windll.kernel32.CreateToolhelp32Snapshot(2, 0)
        if handle == -1:
            return False

        class PROCESSENTRY32(ctypes.Structure):
            _fields_ = [("dwSize", ctypes.c_ulong),
                        ("cntUsage", ctypes.c_ulong),
                        ("th32ProcessID", ctypes.c_ulong),
                        ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
                        ("th32ModuleID", ctypes.c_ulong),
                        ("cntThreads", ctypes.c_ulong),
                        ("th32ParentProcessID", ctypes.c_ulong),
                        ("pcPriClassBase", ctypes.c_long),
                        ("dwFlags", ctypes.c_ulong),
                        ("szExeFile", ctypes.c_char * 260)]

        pe32 = PROCESSENTRY32()
        pe32.dwSize = ctypes.sizeof(PROCESSENTRY32)

        explorer_found = False
        if ctypes.windll.kernel32.Process32First(handle, ctypes.byref(pe32)):
            while True:
                try:
                    exe_name = pe32.szExeFile.decode('gbk', errors='ignore').lower()
                    if exe_name == 'explorer.exe':
                        explorer_found = True
                        break
                except:
                    pass
                if not ctypes.windll.kernel32.Process32Next(handle, ctypes.byref(pe32)):
                    break
        ctypes.windll.kernel32.CloseHandle(handle)
        return explorer_found
    except Exception:
        return True   # 检测失败时默认就绪，避免死等

# 在 亮屏进入桌面.py 中更新 is_desktop_ready 或新增一个更严格的检查

def is_really_unlocked():
    """
    严格检测是否真正解锁进入桌面。
    """
    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        if not hwnd:
            return False
        
        class_name_buff = ctypes.create_unicode_buffer(256)
        ctypes.windll.user32.GetClassNameW(hwnd, class_name_buff, 256)
        class_name = class_name_buff.value.lower()
        
        # 1. 黑名单检查
        lock_classes = [
            "lockapp", "logonui", "windows.ui.core.corewindow", 
            "credentialproviderwrapper", "winlogon"
        ]
        
        if any(cls in class_name for cls in lock_classes):
            return False
            
        # 2. 标题检查
        length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
        if length > 0:
            title_buff = ctypes.create_unicode_buffer(length + 1)
            ctypes.windll.user32.GetWindowTextW(hwnd, title_buff, length + 1)
            title = title_buff.value.lower()
            if "lock" in title or "logon" in title or "sign in" in title:
                return False

        # 3. 【增强】检查任务栏是否存在且可见
        hwnd_taskbar = ctypes.windll.user32.FindWindowW("Shell_TrayWnd", None)
        if hwnd_taskbar:
            # 确保任务栏是可见的，防止资源管理器重启瞬间的假死
            if ctypes.windll.user32.IsWindowVisible(hwnd_taskbar):
                return True
            
        # 如果前台不是锁屏，但任务栏还没出来，可能是过渡状态，返回 False 继续等
        return False

    except Exception:
        return False

def is_explorer_running():
    """单独提取 Explorer 检查逻辑（保留作为备用，但不再作为解锁的唯一标准）"""
    try:
        handle = ctypes.windll.kernel32.CreateToolhelp32Snapshot(2, 0)
        if handle == -1:
            return False
        class PROCESSENTRY32(ctypes.Structure):
            _fields_ = [("dwSize", ctypes.c_ulong), ("cntUsage", ctypes.c_ulong),
                        ("th32ProcessID", ctypes.c_ulong), ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
                        ("th32ModuleID", ctypes.c_ulong), ("cntThreads", ctypes.c_ulong),
                        ("th32ParentProcessID", ctypes.c_ulong), ("pcPriClassBase", ctypes.c_long),
                        ("dwFlags", ctypes.c_ulong), ("szExeFile", ctypes.c_char * 260)]
        pe32 = PROCESSENTRY32()
        pe32.dwSize = ctypes.sizeof(PROCESSENTRY32)
        found = False
        if ctypes.windll.kernel32.Process32First(handle, ctypes.byref(pe32)):
            while True:
                try:
                    if pe32.szExeFile.decode('gbk', errors='ignore').lower() == 'explorer.exe':
                        found = True
                        break
                except:
                    pass
                if not ctypes.windll.kernel32.Process32Next(handle, ctypes.byref(pe32)):
                    break
        ctypes.windll.kernel32.CloseHandle(handle)
        return found
    except:
        return True
def is_lock_screen_active():
    """
    检测当前是否处于锁屏状态
    返回：True = 锁屏中，False = 已解锁
    """
    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        if not hwnd:
            return False
        class_name_buff = ctypes.create_unicode_buffer(256)
        ctypes.windll.user32.GetClassNameW(hwnd, class_name_buff, 256)
        class_name = class_name_buff.value.lower()
        # 锁屏/登录窗口的常见类名
        lock_classes = ["lockapp", "logonui", "windows.ui.core.corewindow"]
        if any(cls in class_name for cls in lock_classes):
            return True
        # 可选：检查窗口标题
        length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
        if length > 0:
            title_buff = ctypes.create_unicode_buffer(length + 1)
            ctypes.windll.user32.GetWindowTextW(hwnd, title_buff, length + 1)
            title = title_buff.value.lower()
            if "lock" in title or "logon" in title:
                return True
        return False
    except Exception:
        return False   # 出错时假设未锁屏

def get_foreground_window():
    """获取前台窗口句柄（辅助检测）"""
    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        return hwnd != 0
    except:
        return True

# ---------- 安全点击重试 ----------
def safe_click(x, y, max_attempts=3):
    """
    安全执行鼠标点击，每次点击前检查锁屏状态，失败后自动重试
    返回 True 表示至少一次点击成功（且解锁状态消失）
    """
    for attempt in range(max_attempts):
        if is_lock_screen_active():
            print(f"  锁屏仍存在，等待后重试 ({attempt+1}/{max_attempts})...")
            time.sleep(0.8)
            continue
        send_mouse_click(x, y, button="left")
        time.sleep(0.5)
        if not is_lock_screen_active():
            print(f"  点击成功（第{attempt+1}次）")
            return True
    print("  警告：多次点击后锁屏仍未解除")
    return False

