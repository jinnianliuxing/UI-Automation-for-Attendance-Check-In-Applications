import uiautomation as auto
import time
import os
from datetime import datetime
import ctypes
import subprocess
import glob
import argparse
import threading
import pythoncom
import json
import sys
import ctypes
from ctypes import wintypes

# 在文件顶部定义全局变量
LOG_CALLBACK = None

def set_log_callback(callback_func):
    """设置日志回调函数"""
    global LOG_CALLBACK
    LOG_CALLBACK = callback_func

def log(msg, force=False):
    """统一日志入口"""
    if VERBOSE_LOG or force:
        try:
            # 1. 保留原有的控制台打印（方便调试）
            print(msg)
            
            # 2. 如果有回调，尝试调用
            if LOG_CALLBACK:
                try:
                    LOG_CALLBACK(msg)
                except Exception as e:
                    # ⚠️ 关键：防止日志回调报错导致打卡中断
                    # 仅在控制台打印回调错误，不中断主流程
                    print(f"[Log Callback Error]: {e}")
                    
        except UnicodeEncodeError:
            print(msg.encode('gbk', errors='replace').decode('gbk'))


# ==================== 全局中断控制 ====================
# 用于主程序通知打卡模块立即停止
CHECKIN_INTERRUPT_EVENT = threading.Event()

def reset_interrupt_flag():
    """重置中断标志，每次开始新任务前调用"""
    CHECKIN_INTERRUPT_EVENT.clear()

def is_interrupted():
    """检查是否被要求中断"""
    return CHECKIN_INTERRUPT_EVENT.is_set()

# ==================== 全局变量 ====================
start_time = time.time()
punch_status = "未知"
punch_message = ""
login_failed = False  # 标记微信登录是否失败
VERBOSE_LOG = True   # 是否启用详细日志（调试阶段建议开启，日常运行可关闭）


# ==================== 路径配置 ====================
USER_HOME = os.path.expanduser("~")
DESKTOP_PATH = os.path.join(USER_HOME, "Desktop")
START_MENU_PATH = os.path.join(USER_HOME, r"AppData\Roaming\Microsoft\Windows\Start Menu\Programs")

WECHAT_PATHS = [
    os.path.join(START_MENU_PATH, "微信.lnk"),
    os.path.join(START_MENU_PATH, "腾讯\微信.lnk"),
    os.path.join(DESKTOP_PATH, "微信.lnk"),
    r"C:\Program Files (x86)\Tencent\WeChat\WeChat.exe",
    r"C:\Program Files\Tencent\WeChat\WeChat.exe",
    os.path.join(USER_HOME, r"AppData\Local\Programs\Tencent\WeChat\WeChat.exe"),
    r"C:\Program Files (x86)\Tencent\Weixin\Weixin.exe",
    r"C:\Program Files\Tencent\Weixin\Weixin.exe",
    os.path.join(USER_HOME, r"AppData\Local\Programs\Tencent\Weixin\Weixin.exe"),
    os.path.join(USER_HOME, r"AppData\Local\Tencent\Weixin\Weixin.exe"),
]

MINI_PROGRAM_KEYWORDS = [
    "中南林业科技大学学生工作部",
    "学生工作部",
]

WAIT_SHORT = 0.3
WAIT_MEDIUM = 0.5
WAIT_LONG = 2.0
RETRY_INTERVAL = 5


# ==================== 窗口管理辅助函数 ====================

def minimize_wechat_windows_gracefully():
    """
    优雅地将微信主窗口置于底层/非激活状态，而不最小化它。
    """
    try:
        root = auto.GetRootControl()
        children = root.GetChildren()
        
        handled_count = 0
        for child in children:
            try:
                if hasattr(child, 'ClassName') and child.ClassName == "mmui::MainWindow":
                    if hasattr(child, 'Name') and "微信" in child.Name:
                        hwnd = child.NativeWindowHandle
                        if hwnd:
                            # SW_SHOWNA (4): 显示但不激活
                            ctypes.windll.user32.ShowWindow(hwnd, 4) 
                            handled_count += 1
            except Exception:
                continue
        
        if handled_count > 0:
            # ✅ 关键：操作完成后，必须等待系统处理完所有窗口消息
            time.sleep(0.8) 
            if VERBOSE_LOG:
                print(f"[窗口管理] 已将 {handled_count} 个微信窗口设为非激活状态")
            
    except Exception as e:
        if VERBOSE_LOG: print(f"[窗口管理] 优雅处理微信窗口失败: {e}")

def force_bring_to_top_retry(ctrl, retries=2):
    """
    强制将控件对应的窗口置顶，并进行多次重试以确保成功。
    增加了更长的等待时间和稳固期。
    """
    if not ctrl:
        return False
    
    hwnd = ctrl.NativeWindowHandle
    if not hwnd:
        return False

    success = False
    for i in range(retries):
        try:
            # 1. 先激活窗口（确保它获得输入焦点权限）
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            time.sleep(0.2) # 等待激活生效
            
            # 2. 再置顶 (HWND_TOPMOST)
            # SWP_NOMOVE (0x0002), SWP_NOSIZE (0x0001), SWP_SHOWWINDOW (0x0040)
            ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0002 | 0x0001 | 0x0040)
            
            # 3. 再次激活（防止置顶过程中焦点丢失）
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            
            time.sleep(0.3) # 每次重试间隔稍长一点，给系统反应时间
            success = True
        except Exception as e:
            if VERBOSE_LOG: print(f"[窗口管理] 置顶尝试 {i+1} 失败: {e}")
            time.sleep(0.5)
    
    if success:
        # 4. 稳固期：置顶成功后，再等待一小会儿，防止其他程序立即抢走焦点
        time.sleep(0.5) 
        if VERBOSE_LOG:
            print(f"[窗口管理] 窗口置顶 ( {retries} 次)")
    return success 
# ==================== 路径工具函数 ====================
def find_valid_path(path_list, file_type="文件"):
    for path in path_list:
        if os.path.exists(path):
            if VERBOSE_LOG: print(f"[路径] 找到{file_type}")
            return path
    return None

def scan_desktop_for_shortcuts(keywords):
    if VERBOSE_LOG: print("[路径] 扫描桌面...")
    try:
        lnk_files = glob.glob(os.path.join(DESKTOP_PATH, "*.lnk"))
        for lnk_file in lnk_files:
            file_name = os.path.basename(lnk_file)
            for keyword in keywords:
                if keyword in file_name:
                    if VERBOSE_LOG: print(f"[路径] 找到快捷方式: {file_name}")
                    return lnk_file
    except Exception as e:
        if VERBOSE_LOG: print(f"[路径] 扫描失败: {e}")
    return None

def scan_start_menu_for_shortcuts(keywords):
    if VERBOSE_LOG: print("[路径] 扫描开始菜单...")
    try:
        lnk_files = glob.glob(os.path.join(START_MENU_PATH, "*.lnk"))
        lnk_files += glob.glob(os.path.join(START_MENU_PATH, "**", "*.lnk"), recursive=True)
        for lnk_file in lnk_files:
            file_name = os.path.basename(lnk_file)
            for keyword in keywords:
                if keyword in file_name:
                    if VERBOSE_LOG: print(f"[路径] 找到快捷方式: {file_name}")
                    return lnk_file
    except Exception as e:
        if VERBOSE_LOG: print(f"[路径] 扫描失败: {e}")
    return None

def get_wechat_path():
    path = find_valid_path(WECHAT_PATHS, "微信")
    if path:
        return path
    
    try:
        import winreg
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Tencent\WeChat", 0, winreg.KEY_READ)
            install_path, _ = winreg.QueryValueEx(key, "InstallPath")
            winreg.CloseKey(key)
            wechat_exe = os.path.join(install_path, "WeChat.exe")
            if os.path.exists(wechat_exe):
                return wechat_exe
        except Exception:
            pass
        
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Tencent\Weixin", 0, winreg.KEY_READ)
            install_path, _ = winreg.QueryValueEx(key, "InstallPath")
            winreg.CloseKey(key)
            weixin_exe = os.path.join(install_path, "Weixin.exe")
            if os.path.exists(weixin_exe):
                return weixin_exe
        except Exception:
            pass
    except Exception as e:
        if VERBOSE_LOG: print(f"[路径] 注册表查找失败: {e}")
    
    return "start wechat:"

def get_mini_program_path():
    """
    增强版小程序路径查找
    """
    
    # ==================== 【第一步】尝试多种位置读取 config.json ====================
    user_custom_path = ""
    
    # 方法 A: 尝试从当前工作目录读取 (通常 exe 运行时 cwd 就是 exe 所在目录)
    config_cwd = os.path.join(os.getcwd(), "config.json")
    
    # 方法 B: 尝试从本脚本所在目录读取 (兼容源码运行)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_script = os.path.join(script_dir, "config.json")
    
    # 方法 C: 尝试从 sys.executable 所在目录读取 (兼容打包后 exe 旁侧的 config)
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        config_exe = os.path.join(exe_dir, "config.json")
    else:
        config_exe = None

    # 按优先级尝试加载
    for conf_path in [config_cwd, config_exe, config_script]:
        if conf_path and os.path.exists(conf_path):
            try:
                with open(conf_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    path = config.get("custom_mini_program_path", "")
                    if path and os.path.exists(path):
                        user_custom_path = path
                        if VERBOSE_LOG: print(f"[路径] 从 {conf_path} 命中用户自定义路径")
                        break # 找到有效的就停止
            except Exception as e:
                if VERBOSE_LOG: print(f"[路径] 读取 {conf_path} 失败: {e}")

    if user_custom_path:
        return user_custom_path
    # ======================================================================

    # ==================== 【第二步】硬编码的备用绝对路径 (可选) ====================
    CUSTOM_ABSOLUTE_PATHS = [
        # r"D:\Users\yangjia\Desktop\中南林业科技大学学生工作部.lnk", 
    ]
    
    for path in CUSTOM_ABSOLUTE_PATHS:
        if os.path.exists(path):
            if VERBOSE_LOG: print(f"[路径] 命中硬编码绝对路径: {path}")
            return path
    # ======================================================================

    # ==================== 【第三步】系统目录扫描 (原有逻辑) ====================
    search_dirs = []
    
    # A. 当前用户桌面和开始菜单
    current_desktop = os.path.join(USER_HOME, "Desktop")
    current_start_menu = os.path.join(USER_HOME, r"AppData\Roaming\Microsoft\Windows\Start Menu\Programs")
    search_dirs.append(("当前用户桌面", current_desktop))
    search_dirs.append(("当前用户开始菜单", current_start_menu))
    
    # B. 公共桌面
    try:
        import ctypes.wintypes
        CSIDL_COMMON_DESKTOPDIRECTORY = 25
        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_COMMON_DESKTOPDIRECTORY, None, 0, buf)
        public_desktop = buf.value
        if os.path.exists(public_desktop):
            search_dirs.append(("公共桌面", public_desktop))
    except Exception:
        pass
        
    # C. 公共开始菜单
    try:
        CSIDL_COMMON_PROGRAMS = 23
        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_COMMON_PROGRAMS, None, 0, buf)
        public_start_menu = buf.value
        if os.path.exists(public_start_menu):
            search_dirs.append(("公共开始菜单", public_start_menu))
    except Exception:
        pass

    # ==================== 【第四步】精确查找 ====================
    for keyword in MINI_PROGRAM_KEYWORDS:
        for dir_name, dir_path in search_dirs:
            target_lnk = os.path.join(dir_path, f"{keyword}.lnk")
            if os.path.exists(target_lnk):
                if VERBOSE_LOG: print(f"[路径] 在 {dir_name} 找到: {target_lnk}")
                return target_lnk

    # ==================== 【第五步】模糊扫描 ====================
    all_scan_paths = [path for _, path in search_dirs]
    
    for scan_path in all_scan_paths:
        if not os.path.exists(scan_path):
            continue
        try:
            lnk_files = glob.glob(os.path.join(scan_path, "*.lnk"))
            if "Start Menu" in scan_path:
                lnk_files += glob.glob(os.path.join(scan_path, "**", "*.lnk"), recursive=True)
                
            for lnk_file in lnk_files:
                file_name = os.path.basename(lnk_file)
                for keyword in MINI_PROGRAM_KEYWORDS:
                    if keyword in file_name:
                        if VERBOSE_LOG: print(f"[路径] 模糊匹配找到 ({scan_path}): {file_name}")
                        return lnk_file
        except Exception as e:
            if VERBOSE_LOG: print(f"[路径] 扫描 {scan_path} 时出错: {e}")

    print("[警告] 未找到小程序快捷方式，请检查桌面或开始菜单")
    return None
def start_application(path):
    try:
        if path.startswith("start "):
            os.system(path)
        elif path.endswith(".lnk"):
            os.startfile(path)
        elif os.path.exists(path):
            subprocess.Popen(path)
        else:
            print(f"[错误] 路径不存在: {path}")
            return False
        return True
    except Exception as e:
        print(f"[错误] 启动异常: {e}")
        return False

def bring_window_to_top(ctrl):
    try:
        hwnd = ctrl.NativeWindowHandle
        if hwnd:
            ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0002 | 0x0001)
            return True
    except Exception:
        pass
    return False

def activate_window(ctrl):
    try:
        hwnd = ctrl.NativeWindowHandle
        if hwnd:
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            return True
    except Exception:
        pass
    return False

def close_app(ctrl, app_name="应用"):
    try:
        if ctrl:
            try:
                if hasattr(ctrl, 'GetWindowPattern'):
                    window_pattern = ctrl.GetWindowPattern()
                    if window_pattern:
                        window_pattern.Close()
                        return
            except Exception:
                pass
            
            try:
                pid = ctrl.ProcessId
                if pid > 0:
                    subprocess.run(
                        ["taskkill", "/F", "/PID", str(pid)],
                        capture_output=True,
                        text=True,
                        encoding='gbk',
                        errors='replace',
                        timeout=5
                    )
            except Exception:
                pass
    except Exception as e:
        if VERBOSE_LOG: print(f"[关闭] {app_name} 出错: {e}")

def record_punch_status(status, message):
    global punch_status, punch_message
    punch_status = status
    punch_message = message
    print(f"[状态] {status}: {message}")

def find_control_by_partial_name(parent, partial_name, max_depth=12):
    def search_recursive(ctrl, current_depth):
        if current_depth > max_depth:
            return None
        if hasattr(ctrl, 'Name') and ctrl.Name and partial_name in ctrl.Name:
            return ctrl
        try:
            for child in ctrl.GetChildren():
                result = search_recursive(child, current_depth + 1)
                if result:
                    return result
        except Exception:
            pass
        return None
    return search_recursive(parent, 0)

def check_and_login_wechat():
    global login_failed
    try:
        log("\n[微信] === 开始检测微信登录状态 ===", force=True)
        
        login_window = auto.WindowControl(Name="微信", ClassName="mmui::LoginWindow", searchDepth=3)
        main_window = auto.WindowControl(Name="微信", ClassName="mmui::MainWindow", searchDepth=5)
        
        # 1. 如果主窗口已存在，说明已登录
        if main_window.Exists(maxSearchSeconds=2):
            log("[微信] 已登录", force=True)
            return True
        
        # 2. 监听登录窗口出现
        log("[微信] 未检测到主窗口，正在监听登录窗口...", force=True)
        login_appeared = False
        for i in range(10): 
            if is_interrupted(): return False
            if login_window.Exists(maxSearchSeconds=0.5) or main_window.Exists(maxSearchSeconds=0.5):
                login_appeared = True
                break
            time.sleep(0.5)

        if not login_appeared:
            log("[微信] 未检测到登录窗口，尝试启动...", force=True)
            wechat_path = get_wechat_path()
            if wechat_path:
                start_application(wechat_path)
            for i in range(15):
                if is_interrupted(): return False
                if login_window.Exists(maxSearchSeconds=0.5) or main_window.Exists(maxSearchSeconds=0.5):
                    login_appeared = True
                    break
                time.sleep(0.5)

        if not login_appeared:
             if main_window.Exists(maxSearchSeconds=2):
                 log("[微信] 启动后已登录", force=True)
                 return True
             print("[错误] 微信启动失败或未响应")
             return False
        
        # 3. 处理冲突弹窗
        log("[微信] 检测到登录窗口，检查是否有冲突提示...", force=True)
        ok_btn = login_window.ButtonControl(Name="我知道了", ClassName="mmui::XOutlineButton", searchDepth=7)
        if not ok_btn.Exists(maxSearchSeconds=2):
            ok_btn = find_control_by_partial_name(login_window, "我知道了", max_depth=7)
            
        if ok_btn and ok_btn.Exists(maxSearchSeconds=2):
            log("[微信] 发现'我知道了'提示，点击处理...", force=True)
            try:
                ok_btn.Click()
                time.sleep(2) 
            except Exception as e:
                log(f"[微信] 点击'我知道了'异常: {e}")
        else:
            log("[微信] 未发现冲突提示，继续正常登录流程")

        # 4. 重新确认窗口状态
        if not login_window.Exists(maxSearchSeconds=3):
             if main_window.Exists(maxSearchSeconds=3):
                 log("[微信] 处理冲突后直接进入主界面", force=True)
                 return True
             else:
                 print("[错误] 处理冲突后登录窗口消失且无主窗口")
                 login_failed = True
                 return False

        # 5. 判断是“进入微信”还是“扫码登录”
        
        # 情况 A: 存在“进入微信”按钮
        login_btn1 = login_window.ButtonControl(Name="进入微信", ClassName="mmui::XOutlineButton", searchDepth=7)
        if login_btn1.Exists(maxSearchSeconds=5):
            log("[微信] 执行自动登录...", force=True)
            try:
                login_btn1.Click()
            except Exception as e:
                log(f"[微信] 点击'进入微信'异常: {e}")
            
            time.sleep(1.0)
            
            log("[微信] 等待登录窗口关闭及主窗口加载...", force=True)
            for i in range(20): 
                if is_interrupted(): return False
                
                if main_window.Exists(maxSearchSeconds=0.5):
                    if not login_window.Exists(maxSearchSeconds=0.5):
                        time.sleep(0.5)
                        log("[微信] 登录成功，主窗口已激活", force=True)
                        return True
                
                # ✅ 【优化点】如果登录窗口还在，尝试更强力的关闭方式
                if login_window.Exists(maxSearchSeconds=0.2):
                    try:
                        if i > 5: 
                             # 方法 1: 尝试 UIA Close
                             lp = login_window.GetWindowPattern()
                             if lp: 
                                 lp.Close()
                                 log("[微信] 尝试 UIA 关闭登录窗口")
                                 time.sleep(0.5)
                                 if not login_window.Exists(maxSearchSeconds=0.5):
                                     continue # 成功关闭
                             
                             # 方法 2: 如果 UIA 失败，使用 Win32 API 发送 WM_CLOSE
                             hwnd = login_window.NativeWindowHandle
                             if hwnd:
                                 log("[微信] UIA 关闭失败，尝试发送 WM_CLOSE 消息...")
                                 WM_CLOSE = 0x0010
                                 ctypes.windll.user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
                                 time.sleep(0.5)
                                 if not login_window.Exists(maxSearchSeconds=0.5):
                                     continue # 成功关闭
                                 
                             # 方法 3: 如果还不行，尝试模拟 Alt+F4
                             log("[微信] 发送 Alt+F4 强制关闭...")
                             unlock_module = __import__('亮屏进入桌面')
                             unlock_module.send_key(0x12, True) # Alt Down
                             unlock_module.send_key(0x73, True) # F4 Down
                             time.sleep(0.1)
                             unlock_module.send_key(0x73, False) # F4 Up
                             unlock_module.send_key(0x12, False) # Alt Up
                             time.sleep(0.5)
                             
                    except Exception as e:
                        log(f"[微信] 关闭登录窗口异常: {e}")
                
                time.sleep(0.5)
            
            print("[错误] 自动登录超时，主窗口未就绪或登录窗口未关闭")
            login_failed = True
            return False
        
        # 情况 B: 存在“扫码登录”提示
        login_btn2 = login_window.TextControl(Name="扫码登录", ClassName="mmui::XTextView", searchDepth=6)
        if login_btn2.Exists(maxSearchSeconds=5):
            print("[警告] 需要扫码登录，请在 120 秒内完成")
            start_wait = time.time()
            scan_success = False
            
            while time.time() - start_wait < 120:
                if is_interrupted(): return False 
                
                if main_window.Exists(maxSearchSeconds=1):
                    if not login_window.Exists(maxSearchSeconds=0.5):
                        log("[微信] 扫码登录成功，主窗口已激活且登录窗已关闭", force=True)
                        scan_success = True
                        break
                    else:
                        # ✅ 【优化点】扫码界面也适用同样的强力关闭逻辑
                        log("[微信] 主窗口已现，但登录窗未退，尝试强制关闭...", force=True)
                        hwnd = login_window.NativeWindowHandle
                        if hwnd:
                            WM_CLOSE = 0x0010
                            ctypes.windll.user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
                            time.sleep(0.5)
                            if not login_window.Exists(maxSearchSeconds=0.5):
                                log("[微信] 强制关闭登录窗成功", force=True)
                                scan_success = True
                                break
                
                time.sleep(1)
            
            if scan_success:
                time.sleep(1.5) 
                return True
            else:
                print("[错误] 扫码登录超时")
                login_failed = True
                return False
        
        # 情况 C: 未知状态
        print("[错误] 未知的微信登录状态")
        login_failed = True
        return False
        
    except Exception as e:
        print(f"[错误] 微信登录异常: {e}")
        login_failed = True
        return False

def handle_wechat_conflict(login_window_ctrl):
    """
    处理微信登录冲突：
    1. 检测是否有“该账号已在...”的提示弹窗
    2. 点击“我知道了”
    3. 关闭随后出现的扫码登录窗口
    """
    log("[冲突处理] 检查是否有账号冲突提示...")
    
    # 常见的冲突提示按钮文本可能包含“我知道了”、“确定”等
    # 这里尝试查找登录窗口下的所有按钮，看是否有“我知道了”
    ok_btn = login_window_ctrl.ButtonControl(Name="我知道了", ClassName="mmui::XOutlineButton", searchDepth=7)
    if not ok_btn.Exists(maxSearchSeconds=2):
        # 尝试模糊匹配
        ok_btn = find_control_by_partial_name(login_window_ctrl, "我知道了", max_depth=7)
        
    if ok_btn and ok_btn.Exists(maxSearchSeconds=2):
        log("[冲突处理] 点击'我知道了'...")
        ok_btn.Click()
        time.sleep(1.5) # 等待扫码窗口弹出
        
        # 此时通常会弹出一个新的扫码登录窗口，或者原来的登录窗口变成扫码状态
        # 我们的策略是：直接关闭当前的登录/扫码窗口，因为主进程其实已经在线了
        log("[冲突处理] 关闭登录/扫码窗口以回归主界面...")
        try:
            login_window_ctrl.GetWindowPattern().Close()
        except Exception:
            pass
        time.sleep(1)
    else:
        log("[冲突处理] 未检测到标准冲突提示，尝试直接关闭登录窗...")
        try:
            login_window_ctrl.GetWindowPattern().Close()
        except Exception:
            pass
        time.sleep(1)

def execute_punch_logic(punch_window):
    """执行打卡核心逻辑 - 精简版"""
    try:
        log("[打卡] 开始分析页面元素...", force=True) # ✅ 新增
        
        dkjl = punch_window.TextControl(Name="打卡记录", searchDepth=16)
        zfw = punch_window.TextControl(Name="您已在打卡范围内", searchDepth=13)
        bzfw = punch_window.TextControl(Name="不在打卡范围内", searchDepth=12)
        wddksj = punch_window.TextControl(Name="未到打卡时间", searchDepth=12)
        dkwc = punch_window.TextControl(Name="打卡完成", searchDepth=12)
        
        my_tab = punch_window.TextControl(Name="我的", searchDepth=7)
        click_login = punch_window.TextControl(Name="点击登录", searchDepth=13)
        
        # 1. 判断是否已在二级页面
        if auto.WaitForExist(dkjl, timeout=5):
            log("[打卡] 检测到'打卡记录'，确认已在二级页面") # ✅ 优化日志
        else:
            log("[打卡] 未在二级页面，在首页...") # ✅ 新增
            # 2. 返回首页并点击打卡
            mainpage = punch_window.TextControl(Name="首页", searchDepth=8)
            if not auto.WaitForExist(mainpage, timeout=5):
                return (False, True, "失败", "首页未加载")
            
            # 查找打卡按钮
            log("[打卡] 正在搜索'季学期平安打卡'按钮...") # ✅ 新增
            main_btn = None
            for i in range(3): 
                main_btn = find_control_by_partial_name(punch_window, "季学期平安打卡", max_depth=12)
                if main_btn: 
                    log(f"[打卡] 第{i+1}次尝试找到按钮") # ✅ 新增
                    break
                time.sleep(1)
                
            if not main_btn:
                # 尝试登录修复
                log("[打卡] 未找到按钮，检查是否需登录...") # ✅ 优化日志
                if my_tab.Exists(maxSearchSeconds=3):
                    my_tab.Click()
                    time.sleep(1)
                    if click_login.Exists(maxSearchSeconds=3):
                        log("[登录] 发现'点击登录'，执行小程序内登录...") # ✅ 优化日志
                        click_login.Click()
                        
                        login_success = False
                        for i in range(6):
                            if find_control_by_partial_name(punch_window, "季学期平安打卡", max_depth=12):
                                login_success = True
                                main_btn = find_control_by_partial_name(punch_window, "季学期平安打卡", max_depth=12)
                                break
                            log(f"[登录] 等待登录生效... ({i+1}/6)") # ✅ 新增
                            time.sleep(5)
                        
                        if not login_success:
                            return (False, True, "失败", "登录超时")
                    else:
                         return (False, True, "需重启", "已登录但未找到按钮")
                else:
                    return (False, True, "失败", "未找到'我的'按钮")
            
            if main_btn:
                log("[打卡] 点击*季学期平安打卡...") # ✅ 新增
                try:
                    main_btn.Click()
                except Exception:
                    parent = main_btn.GetParentControl()
                    parent.Click()
            else:
                return (False, True, "失败", "*季学期平安打卡未找到")

            if not auto.WaitForExist(dkjl, timeout=5):
                return (False, True, "失败", "二级页面未出现")
        
        # 执行打卡
        log("[打卡] 进入状态判断流程...") # ✅ 新增
        if auto.WaitForExist(zfw, timeout=5):
            log("[打卡] 状态：在打卡范围内，寻找'平安打卡'按钮") # ✅ 优化日志
            pingan_btn = None
            current = zfw
            # 向上查找父容器中的“平安打卡”按钮
            for level in range(2):
                parent = current.GetParentControl()
                btn = parent.TextControl(Name="平安打卡", searchDepth=3)
                if btn.Exists(maxSearchSeconds=2):
                    pingan_btn = btn
                    break
                current = parent
            
            if pingan_btn:
                log("[打卡] 找到'平安打卡'按钮，执行点击") # ✅ 新增
                try:
                    pingan_btn.Click()
                    if auto.WaitForExist(dkwc, timeout=8):
                        now = datetime.now()
                        return (True, False, "完成", f"打卡成功 {now.strftime('%H:%M:%S')}")
                    else:
                        return (False, True, "失败", "未检测到完成界面")
                except Exception as e:
                    return (False, True, "失败", f"点击异常: {str(e)}")
            else:
                return (False, True, "失败", "未找到平安打卡按钮")

        elif auto.WaitForExist(bzfw, timeout=5):
            log("[打卡] 状态：不在打卡范围内") # ✅ 新增
            return (False, True, "不在范围", f"不在范围内 {datetime.now().strftime('%H:%M:%S')}")

        elif auto.WaitForExist(dkwc, timeout=5):
            log("[打卡] 状态：已完成打卡") # ✅ 新增
            return (True, False, "完成", f"已打卡 {datetime.now().strftime('%H:%M:%S')}")

        elif auto.WaitForExist(wddksj, timeout=5):
            log("[打卡] 状态：未到打卡时间") # ✅ 新增
            return (False, False, "未到时间", f"未到打卡时间 {datetime.now().strftime('%H:%M:%S')}")

        else:
            log("[打卡] 状态：未知，未匹配到任何已知文本") # ✅ 新增
            return (False, True, "未知", "未知状态")

    except Exception as e:
        print(f"[错误] 打卡执行异常: {e}")
        return (False, True, "失败", f"执行异常: {str(e)}")

def run_punch_task(punch_window):
    if punch_window:
        success, should_retry, status, message = execute_punch_logic(punch_window)
        record_punch_status(status, message)
        close_app(punch_window, "打卡小程序")
        return success, should_retry
    return False, False

def parse_args():
    parser = argparse.ArgumentParser(description='自动打卡并发送微信通知')
    parser.add_argument('--notify', type=str, default='0', help='是否发送微信通知 (1=发送, 0=不发送)')
    parser.add_argument('--close-wechat', type=str, default='1', help='发送通知后是否关闭微信 (1=关闭, 0=保持打开)')
    return parser.parse_args()

# # ==================== 主流程 ====================
# print("=" * 30)
# print("[开始] 执行打卡任务")
# print("=" * 30)

# punch_start = time.time()
# punch_window = None
# max_retry = 3
# login_processed = False
# task_completed = False

# mini_program_path = get_mini_program_path()

# for retry in range(max_retry):
#     try:
#         log(f"\n[尝试 {retry + 1}/{max_retry}] 启动小程序...", force=True)
        
#         if mini_program_path:
#             start_application(mini_program_path)
#         else:
#             print("[错误] 小程序路径未配置")
#             record_punch_status("失败", "路径未配置")
#             break
        
#         punch_window = auto.PaneControl(Name="中南林业科技大学学生工作部", ClassName="Chrome_WidgetWin_0", searchDepth=3)
#         if auto.WaitForExist(punch_window, timeout=10):
#             log("[窗口] 主窗口已出现")
#             bring_window_to_top(punch_window)
#             activate_window(punch_window)
#             time.sleep(0.3)
            
#             success, should_retry = run_punch_task(punch_window)
#             punch_window = None
            
#             if success:
#                 log("[结果] 打卡成功", force=True)
#                 task_completed = True
#                 break
#             elif not should_retry:
#                 log(f"[结果] {punch_status}", force=True)
#                 task_completed = True
#                 break
#             else:
#                 if retry < max_retry - 1:
#                     log(f"[重试] 失败 ({punch_status})，{RETRY_INTERVAL}秒后重试...")
#                     time.sleep(RETRY_INTERVAL)
#                     continue
#                 else:
#                     print("[错误] 已达最大重试次数")
#                     break
#         else:
#             log("[窗口] 主窗口未出现")
#             if login_failed:
#                 print("[错误] 登录已失败，终止")
#                 break
            
#             if not login_processed:
#                 log("[处理] 检测微信登录...")
#                 login_result = check_and_login_wechat()
                
#                 if login_result:
#                     log("[处理] 登录完成，重试启动...")
#                     login_processed = True
#                     time.sleep(1)
#                     continue
#                 else:
#                     print("[错误] 微信登录处理失败")
#                     break
            
#             record_punch_status("失败", "小程序启动失败")
#             break

#     except Exception as e:
#         print(f"[错误] 尝试 {retry + 1} 异常: {e}")
#         if punch_window:
#             close_app(punch_window, "打卡小程序")
#             punch_window = None
        
#         if not login_processed and not login_failed:
#             log("[处理] 异常后尝试登录...")
#             login_result = check_and_login_wechat()
#             if login_result:
#                 login_processed = True
#                 time.sleep(1)
#                 continue
#             else:
#                 break
#         else:
#             if retry < max_retry - 1:
#                 time.sleep(RETRY_INTERVAL)
#                 continue
#             else:
#                 record_punch_status("失败", f"程序异常: {str(e)}")
#                 break

# punch_end = time.time()
# log(f"[耗时] 打卡阶段: {punch_end - punch_start:.2f}s", force=True)

# # ==================== 微信操作流程 ====================
# print("\n" + "=" * 30)
# print("[开始] 微信操作")
# print("=" * 30)

# args = parse_args()
# enable_notify = args.notify == '1'
# close_wechat = args.close_wechat == '1'

# # 【新增】打印接收到的参数，方便排查 GUI 传递问题
# print(f"[DEBUG] 接收到的参数: notify={args.notify}, close_wechat={args.close_wechat}")
# print(f"[DEBUG] 解析结果: enable_notify={enable_notify}, close_wechat={close_wechat}")

# wechat_start = time.time()
# wechat_window = None

# try:
#     need_wechat_operation = enable_notify or close_wechat
#     log(f"[调试] need_wechat_operation={need_wechat_operation}, login_failed={login_failed}", force=True)
    
#     if not need_wechat_operation:
#         log("[微信] 无需操作", force=True)
#     elif login_failed:
#         print("[警告] 之前微信登录已失败，跳过通知")
#     else:
#         log("[微信] 进入微信操作主流程", force=True)
#         # --- [步骤 1] 查找现有的微信主窗口 ---
#         log("[微信] 查找已运行的微信主窗口...", force=True)
#         main_win = auto.WindowControl(ClassName="mmui::MainWindow", searchDepth=5)
        
#         # 等待一小会儿看主窗口是否直接存在
#         log("[调试] 检查 main_win.Exists...", force=True)
#         if main_win.Exists(maxSearchSeconds=3):
#             wechat_window = main_win
#             log("[微信] 发现主窗口，状态正常", force=True)
#         else:
#             log("[微信] 未找到主窗口，检查登录状态...", force=True)
#             # --- [步骤 2] 主窗口不存在，检查是否有登录弹窗 ---
#             login_win = auto.WindowControl(ClassName="mmui::LoginWindow", searchDepth=3)
            
#             if login_win.Exists(maxSearchSeconds=5):
#                 log("[微信] 检测到登录窗口，尝试恢复会话...", force=True)
#                 handle_wechat_conflict(login_win) 
                
#                 login_win = auto.WindowControl(ClassName="mmui::LoginWindow", searchDepth=3)
                
#                 if login_win.Exists(maxSearchSeconds=3):
#                     enter_btn = login_win.ButtonControl(Name="进入微信", ClassName="mmui::XOutlineButton")
#                     if enter_btn.Exists(maxSearchSeconds=3):
#                         log("[微信] 点击'进入微信'...", force=True)
#                         try:
#                             enter_btn.Click()
#                         except Exception as e:
#                             log(f"[微信] 点击按钮异常: {e}")
                        
#                         time.sleep(2)
                        
#                         if auto.WaitForExist(main_win, timeout=10):
#                             wechat_window = main_win
#                             log("[微信] 登录恢复成功", force=True)
#                         else:
#                             log("[微信] 登录后未出现主窗口，再次检查冲突...", force=True)
#                             current_login = auto.WindowControl(Name="微信", ClassName="mmui::LoginWindow", searchDepth=3)
#                             if current_login.Exists():
#                                 handle_wechat_conflict(current_login)
                            
#                             if auto.WaitForExist(main_win, timeout=5):
#                                 wechat_window = main_win
#                             else:
#                                 raise Exception("登录后仍未获取到主窗口")
#                     else:
#                         log("[微信] 未发现'进入微信'按钮，处于扫码或其他状态...", force=True)
#                         handle_wechat_conflict(login_win)
#                         time.sleep(2)
#                         if main_win.Exists(maxSearchSeconds=5):
#                             wechat_window = main_win
#                         else:
#                             raise Exception("微信处于异常登录状态，请人工干预")
#                 else:
#                      raise Exception("登录窗口在处理冲突后消失，状态异常")
#             else:
#                 log("[微信] 未发现任何微信窗口，尝试启动...", force=True)
#                 # --- [步骤 3] 既无主窗口也无登录窗，尝试温和启动 ---
#                 wechat_path = get_wechat_path()
#                 if wechat_path and wechat_path != "start wechat:":
#                     start_application(wechat_path)
#                     log("[微信] 已启动微信，等待主窗口...", force=True)
#                     if auto.WaitForExist(main_win, timeout=15):
#                         wechat_window = main_win
#                     else:
#                         login_win = auto.WindowControl(Name="微信", ClassName="mmui::LoginWindow", searchDepth=3)
#                         if login_win.Exists(maxSearchSeconds=5):
#                              log("[微信] 启动后出现登录窗，尝试处理...", force=True)
#                              handle_wechat_conflict(login_win)
#                              if auto.WaitForExist(main_win, timeout=5):
#                                  wechat_window = main_win
#                              else:
#                                  raise Exception("微信启动后卡在登录界面")
#                         else:
#                             raise Exception("微信启动超时且未找到窗口")
#                 else:
#                     raise Exception("未找到微信安装路径")

#         # --- [步骤 4] 最终校验 ---
#         log(f"[调试] 最终校验 wechat_window: {wechat_window}", force=True)
#         if not wechat_window or not wechat_window.Exists(maxSearchSeconds=2):
#             raise Exception("无法获取有效的微信主窗口句柄")
        
#         log("[微信] 准备发送消息...", force=True)

#         # --- [步骤 5] 执行发送逻辑 (参照旧版逻辑优化) ---
#         if enable_notify:
#             log("[微信] 开始执行发送逻辑...", force=True)
#             bring_window_to_top(wechat_window)
#             activate_window(wechat_window)
#             time.sleep(0.5) 
            
#             # 防御性检查：防止在激活瞬间弹出登录窗
#             login_check = auto.WindowControl(ClassName="mmui::LoginWindow", searchDepth=3)
#             if login_check.Exists(maxSearchSeconds=1):
#                 log("[微信] 意外弹出登录窗，尝试关闭...", force=True)
#                 try: login_check.GetWindowPattern().Close()
#                 except: pass
#                 time.sleep(1)
#                 bring_window_to_top(wechat_window)
#                 activate_window(wechat_window)

#             # 1. 定位搜索框并输入
#             search = wechat_window.EditControl(ClassName="mmui::XValidatorTextEdit")
#             if auto.WaitForExist(search, timeout=5):
#                 log("[微信] 找到搜索框，输入'文件传输助手'...", force=True)
#                 search.Click()
#                 time.sleep(0.3)
#                 search.SendKeys("文件传输助手")
#                 time.sleep(0.3)
#                 search.SendKeys("{Enter}")
                
#                 # 2. 等待搜索结果出现并进入聊天窗口
#                 # 旧版逻辑：等待 TextControl 出现。这里我们稍微增加一点容错时间
#                 time.sleep(1.5) 
                


#                 # 3. 定位聊天输入框 (关键修复：使用 AutomationId 和足够深的搜索层级)
#                 log("[微信] 正在定位聊天输入框 (AutomationId: chat_input_field)...", force=True)
                
#                 # 方案 A: 使用 AutomationId (最稳，无视层级)
#                 chatipt = wechat_window.EditControl(AutomationId="chat_input_field", searchDepth=14)
                
#                 # 方案 B: 如果方案 A 失败，尝试 ClassName + Name (兼容旧版)
#                 if not chatipt.Exists(maxSearchSeconds=1):
#                     chatipt = wechat_window.EditControl(Name="文件传输助手", ClassName="mmui::ChatInputField")
                
#                 if chatipt.Exists(maxSearchSeconds=5):
#                     log("[微信] 成功定位输入框，准备发送消息", force=True)
                    
#                     # 构造消息内容
#                     final_message = punch_message
#                     if not final_message or not final_message.strip():
#                         if punch_status and punch_status != "未知":
#                             final_message = f"打卡状态: {punch_status}"
#                         else:
#                             final_message = "打卡脚本执行完毕。"
                    
#                     if final_message:
#                         log(f"[微信] 发送内容: {final_message}", force=True)
#                         chatipt.SendKeys(final_message)
#                         time.sleep(0.3)
#                         chatipt.SendKeys("{Enter}")
#                         log(f"[微信] 消息已发送", force=True)
#                     else:
#                         log("[微信] 消息内容为空，跳过发送", force=True)
#                 else:
#                     log("[错误] 未找到聊天输入框 (请检查微信版本或窗口层级)", force=True)
#             else:
#                 log("[错误] 未找到微信搜索框", force=True)
            
#             time.sleep(0.5)
#         else:
#             log("[微信] enable_notify 为 False，跳过发送", force=True)
        
#         # --- [步骤 6] 关闭逻辑 ---
#         if close_wechat:
#             log("[微信] 关闭微信...", force=True)
#             time.sleep(2)
#             try:
#                 if hasattr(wechat_window, 'GetWindowPattern'):
#                      wechat_window.GetWindowPattern().Close()
#                 else:
#                      raise Exception("No WindowPattern")
#             except Exception:
#                 if wechat_window.ProcessId > 0:
#                     subprocess.run(
#                         ["taskkill", "/F", "/PID", str(wechat_window.ProcessId)],
#                         capture_output=True,
#                         text=True,
#                         encoding='gbk',
#                         errors='replace'
#                     )
#             log("[微信] 已关闭", force=True)
#         else:
#             if enable_notify:
#                 log("[微信] 保持打开")

#         wechat_end = time.time()
#         log(f"[耗时] 微信阶段: {wechat_end - wechat_start:.2f}s", force=True)

# except Exception as e:
#     print(f"[错误] 微信操作异常: {e}")
#     import traceback
#     traceback.print_exc()




# # ==================== 汇总 ====================
# print("\n" + "=" * 30)
# print("[汇总]")
# print("=" * 30)
# print(f"状态: {punch_status}")
# print(f"消息: {punch_message}")
# print(f"总耗时: {time.time() - start_time:.2f} 秒")
# print("=" * 30)
# 【新增】统一入口函数，供主程序调用
def run_full_checkin_task(enable_notify_str="0", close_wechat_str="1"):
    """
    执行完整的打卡+通知流程
    参数:
        enable_notify_str: "1" 或 "0"
        close_wechat_str: "1" 或 "0"
    返回:
        (success: bool, status: str, message: str)
    """
    # ✅ 关键修复：初始化当前线程的 COM 环境
    try:
        pythoncom.CoInitializeEx(0x2) 
    except Exception:
        pass

    global punch_status, punch_message, login_failed
    
    try:
        reset_interrupt_flag()
        
        punch_status = "未知"
        punch_message = ""
        login_failed = False
        
        enable_notify = enable_notify_str == '1'
        close_wechat = close_wechat_str == '1'

        punch_start = time.time()
        punch_window = None
        max_retry = 4 
        login_processed = False # 标记是否已经处理过微信登录状态
        task_completed = False

        mini_program_path = get_mini_program_path()

        for retry in range(max_retry):
            if is_interrupted():
                record_punch_status("中断", "用户手动终止")
                return False, "中断", "用户手动终止"

            try:
                log(f"\n[尝试 {retry + 1}/{max_retry}] 启动小程序...", force=True)
                
                if mini_program_path:
                    # ✅ 优化分支：根据是否刚登录过，调整“微信置底”的时机
                    
                    if not login_processed:
                        # 情况 A：正常启动或非登录后重试
                        # 策略：先置底微信，防止遮挡新启动的小程序
                        minimize_wechat_windows_gracefully()
                        start_application(mini_program_path)
                        time.sleep(1.0) # 正常启动等待
                    else:
                        # 情况 B：刚完成微信登录后的首次重试
                        # 策略：先启动小程序，让其自然获取焦点，稍后再处理微信窗口
                        log("[策略] 登录后首次启动，暂不干扰微信焦点，直接启动小程序...")
                        start_application(mini_program_path)
                        time.sleep(2.0) # 登录后启动稍慢，多等1秒
                        
                        # ✅ 关键修改：在启动后、检测前，再将微信置底
                        # 此时小程序应该已经弹出，将微信置底可以防止后续操作被微信抢焦点
                        minimize_wechat_windows_gracefully()

                else:
                    record_punch_status("失败", "小程序路径未配置，请手动选择路径")
                    break
                
                # 3. ✅ 简化版：直接等待窗口出现 (最多3秒)
                punch_window_found = False
                possible_classnames = ["Chrome_WidgetWin_0", "Chrome_WidgetWin_1"]
                
                log("[检测] 正在查找小程序窗口...")
                
                # 使用库自带的 WaitForExist，比手动 for 循环更简洁
                # 我们尝试两种类名，只要找到一个即可
                for cname in possible_classnames:
                    temp_ctrl = auto.PaneControl(Name="中南林业科技大学学生工作部", ClassName=cname, searchDepth=5)
                    # timeout=3 表示最多等3秒
                    if auto.WaitForExist(temp_ctrl, timeout=3):
                        punch_window = temp_ctrl
                        punch_window_found = True
                        break
                
                if punch_window_found:
                    if is_interrupted():
                        close_app(punch_window, "打卡小程序")
                        record_punch_status("中断", "用户手动终止")
                        return False, "中断", "用户手动终止"

                    log("[窗口] 小程序已出现，执行置顶...")
                    time.sleep(0.3)
                    
                    # ✅ 优化：置顶重试次数减少，速度加快
                    force_bring_to_top_retry(punch_window, retries=1) 
                    activate_window(punch_window)
                    
                    # ==================== 【核心修复】新增：强制激活焦点 ====================
                    log("[窗口] 正在强制激活小程序焦点...")
                    
                    # 1. 再次强制置顶，确保它在最上层
                    hwnd = punch_window.NativeWindowHandle
                    if hwnd:
                        # SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
                        ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0002 | 0x0001 | 0x0040)
                        # 再次设置为前台窗口
                        ctypes.windll.user32.SetForegroundWindow(hwnd)
                    
                    # 2. 关键：模拟一次微小的鼠标移动或点击，以“唤醒”窗口的输入响应
                    # 有些小程序窗口需要真实的鼠标事件才能激活 UIA 树
                    try:
                        # 获取窗口中心坐标
                        rect = punch_window.BoundingRectangle
                        center_x = (rect.left + rect.right) // 2
                        center_y = (rect.top + rect.bottom) // 2
                        
                        # 移动鼠标到窗口中心 (这会强制 Windows 将该窗口设为活动窗口)
                        ctypes.windll.user32.SetCursorPos(center_x, center_y)
                        time.sleep(0.2)
                        
                        # 可选：如果还是不稳定，可以取消下面这行的注释，模拟一次左键点击
                        # send_mouse_click(center_x, center_y, button="left") 
                        # time.sleep(0.2)
                        
                    except Exception as e:
                        log(f"[窗口] 鼠标激活异常: {e}")

                    time.sleep(0.8) # 给系统一点时间处理焦点切换和 UI 渲染
                    # ========================================================================

                    success, should_retry = run_punch_task(punch_window)
                    punch_window = None
                    
                    if success:
                        log("[结果] 打卡成功", force=True)
                        task_completed = True
                        break
                    elif not should_retry:
                        log(f"[结果] {punch_status}", force=True)
                        task_completed = True
                        break
                    else:
                        # 打卡逻辑内部失败（如按钮找不到），需要重试
                        if retry < max_retry - 1:
                            log(f"[重试] 打卡逻辑失败 ({punch_status})，{RETRY_INTERVAL}秒后重试...")
                            for _ in range(6): 
                                if is_interrupted():
                                    record_punch_status("中断", "用户手动终止")
                                    return False, "中断", "用户手动终止"
                                time.sleep(0.5)
                            continue
                        else:
                            break
                else:
                    # --- 核心优化分支：窗口未出现 (保留你要求的流程) ---
                    log("[窗口] 主窗口未出现")
                    
                    # ✅ 核心优化：如果窗口没出现，先快速检查微信是否真的活着
                    # 避免直接进入耗时的 check_and_login_wechat
                    main_win_check = auto.WindowControl(ClassName="mmui::MainWindow", searchDepth=5)
                    wechat_alive = main_win_check.Exists(maxSearchSeconds=1)
                    
                    if wechat_alive:
                        log("[诊断] 微信主窗口存在，但小程序未启动。可能是快捷方式无效或小程序卡死。")
                        # 如果微信活着但小程序没起来，通常重试也没用，除非重启微信
                        # 这里选择直接标记失败，或者尝试杀进程重启（可选）
                        if retry < max_retry - 1:
                             log("[策略] 尝试强制关闭残留小程序进程并重试...")
                             subprocess.run(["taskkill", "/F", "/IM", "WeChatAppEx.exe"], creationflags=0x08000000, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                             time.sleep(1)
                             continue
                        else:
                            record_punch_status("失败", "小程序启动超时，请检查快捷方式")
                            break
                    else:
                        # 微信也不在，才去执行登录流程
                        if not login_processed:
                            log("[处理] 微信未运行，开始登录流程...")
                            login_result = check_and_login_wechat()
                            
                            if login_result:
                                log("[处理] 登录完成，重试启动...")
                                login_processed = True
                                # ✅ 注意：这里不再调用 minimize_wechat_windows_gracefully()
                                # 而是让循环回到开头，由上面的 if not login_processed 分支处理
                                continue 
                            else:
                                log("[错误] 微信登录失败")
                                break
                        else:
                            # 已经处理过登录了，还是起不来，放弃
                            record_punch_status("失败", "小程序启动失败")
                            break

            except Exception as e:
                if is_interrupted():
                    record_punch_status("中断", "用户手动终止")
                    return False, "中断", "用户手动终止"

                log(f"[错误] 尝试 {retry + 1} 异常: {e}")
                if punch_window:
                    close_app(punch_window, "打卡小程序")
                    punch_window = None
                
                # 异常后的重试逻辑保持不变，但建议缩短等待
                if retry < max_retry - 1:
                    time.sleep(2) 
                    continue
                else:
                    record_punch_status("失败", f"程序异常: {str(e)}")
                    break

        # --- 微信通知部分 (保持不变，但确保快速退出) ---
        if is_interrupted():
             record_punch_status("中断", "用户手动终止")
             return False, "中断", "用户手动终止"

        wechat_start = time.time()
        wechat_window = None
        
        need_wechat_operation = enable_notify or close_wechat
        
        if need_wechat_operation and not login_failed:
            try:
                main_win = auto.WindowControl(ClassName="mmui::MainWindow", searchDepth=5)
                if main_win.Exists(maxSearchSeconds=2): # 缩短等待
                    wechat_window = main_win
                else:
                    # 如果之前登录过，这里应该能直接找到，找不到说明微信挂了
                    pass 
                
                if wechat_window and wechat_window.Exists():
                    if enable_notify:
                        # ... (发送逻辑保持不变，但可以简化日志)
                        bring_window_to_top(wechat_window)
                        activate_window(wechat_window)
                        time.sleep(0.3)
                        
                        search = wechat_window.EditControl(ClassName="mmui::XValidatorTextEdit")
                        if search.Exists(maxSearchSeconds=2): # 缩短等待
                            search.Click()
                            time.sleep(0.2)
                            search.SendKeys("文件传输助手")
                            time.sleep(0.2)
                            search.SendKeys("{Enter}")
                            time.sleep(1.0)
                            
                            chatipt = wechat_window.EditControl(AutomationId="chat_input_field", searchDepth=14)
                            if not chatipt.Exists(maxSearchSeconds=1):
                                chatipt = wechat_window.EditControl(Name="文件传输助手", ClassName="mmui::ChatInputField")
                            
                            if chatipt.Exists(maxSearchSeconds=2): # 缩短等待
                                final_message = punch_message if punch_message else f"打卡状态: {punch_status}"
                                chatipt.SendKeys(final_message)
                                time.sleep(0.2)
                                chatipt.SendKeys("{Enter}")
                                log("[微信] 消息已发送")
                    
                    if close_wechat:
                        time.sleep(1)
                        try:
                            wechat_window.GetWindowPattern().Close()
                        except:
                            if wechat_window.ProcessId > 0:
                                subprocess.run(["taskkill", "/F", "/PID", str(wechat_window.ProcessId)], capture_output=True)
            except Exception as e:
                if not is_interrupted():
                    log(f"[微信] 操作异常: {e}")

        if is_interrupted():
            return False, "中断", "用户手动终止"
            
        return (punch_status == "完成" or punch_status == "成功"), punch_status, punch_message

    except Exception as e:
        if is_interrupted():
            return False, "中断", "用户手动终止"
        return False, "异常", str(e)
    
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass