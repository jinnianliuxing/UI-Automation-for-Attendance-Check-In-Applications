# -*- coding: utf-8 -*-

# ==================== DPI 适配修复 (解决窗口过小/模糊) ====================
import ctypes
import sys
import os

# 1. 设置高 DPI 感知 (必须在创建 Tk 窗口之前调用)
try:
    # Windows 8.1+
    ctypes.windll.shcore.SetProcessDpiAwareness(2) 
except Exception:
    try:
        # Windows Vista+
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# 2. 获取当前屏幕的 DPI 缩放比例
def get_dpi_scale():
    """获取当前主显示器的 DPI 缩放比例 (例如 1.0, 1.25, 1.5)"""
    try:
        # 获取 HDC
        hdc = ctypes.windll.user32.GetDC(0)
        # 获取逻辑 DPI (通常返回 96, 120, 144 等)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88) # LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(0, hdc)
        return dpi / 96.0
    except Exception:
        return 1.0

# 获取缩放比例
DPI_SCALE = get_dpi_scale()

# ==================== 正常导入模块 ====================
import pystray
from PIL import Image, ImageDraw
import win32com.client
import win32event
import win32api
import winerror
import win32gui
import win32process  
import win32con
from ctypes import wintypes
import uiautomation as auto
current_dir = os.path.dirname(os.path.abspath(__file__))

# 如果当前目录不在 sys.path 中，则添加
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# 然后再导入
import 亮屏进入桌面 as unlock_module
import 亮屏进入桌面 as unlock_module
import 打卡并发消息 as punch_module 
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from tkinter import filedialog
from datetime import datetime, timedelta
import subprocess
import re
import json
import traceback


# # ========== 嵌入式 Python Tkinter 环境配置 ==========
# base_dir = os.path.dirname(sys.executable)          # Main File 目录

# # 1. 添加标准库路径（关键！）
# lib_dir = os.path.join(base_dir, 'Lib')
# if os.path.exists(lib_dir):
#     sys.path.append(lib_dir)                        # 让 Python 能找到 tkinter 包

# # 2. 添加 DLLs 目录到模块搜索路径和 DLL 搜索路径
# dlls_dir = os.path.join(base_dir, 'DLLs')
# if os.path.exists(dlls_dir):
#     sys.path.append(dlls_dir)                       # 让 Python 能找到 _tkinter.pyd
#     if hasattr(os, 'add_dll_directory'):
#         os.add_dll_directory(dlls_dir)              # Python 3.8+ 必需
#     else:
#         os.environ['PATH'] = dlls_dir + os.pathsep + os.environ.get('PATH', '')

# # 3. 设置 Tcl/Tk 脚本路径
# tcl_dir = os.path.join(base_dir, 'tcl')
# tcl_lib = os.path.join(tcl_dir, 'tcl8.6')
# tk_lib = os.path.join(tcl_dir, 'tk8.6')
# if os.path.exists(tcl_lib):
#     os.environ['TCL_LIBRARY'] = tcl_lib
# if os.path.exists(tk_lib):
#     os.environ['TK_LIBRARY'] = tk_lib




# 获取程序运行的基础目录
if getattr(sys, 'frozen', False):
    # 如果是打包后的 exe
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    # 如果是源码运行
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")

# ========== 窗口尺寸配置 ==========
# 基础设计尺寸 (基于 100% 缩放)
BASE_WIDTH_RAW = 450
BASE_HEIGHT_RAW = 570

# 根据 DPI 自动调整尺寸
BASE_WIDTH = int(BASE_WIDTH_RAW * DPI_SCALE)
BASE_HEIGHT = int(BASE_HEIGHT_RAW * DPI_SCALE)
FOLD_RATIO = 0.65
# ===============================================
# 全局常量：计划任务名称（主界面和二级窗口共享）
SCHEDULED_TASK_NAME = "自动平安打卡"

# ==================== API 防睡眠工具类 ====================
class ApiSleepPreventer:
    """
    基于 Windows API SetThreadExecutionState 的防睡眠工具。
    策略：允许屏幕关闭，但禁止系统进入睡眠或休眠。
    """
    def __init__(self):
        self.ES_CONTINUOUS = 0x80000000
        self.ES_SYSTEM_REQUIRED = 0x00000001  # 关键：阻止系统睡眠/休眠
        # 注意：这里故意不包含 ES_DISPLAY_REQUIRED，所以屏幕可以正常关闭
        self.is_preventing = False

    def prevent(self):
        """开启防睡眠（允许屏保/黑屏）"""
        if not self.is_preventing:
            # 只设置 SYSTEM_REQUIRED，不设置 DISPLAY_REQUIRED
            ctypes.windll.kernel32.SetThreadExecutionState(
                self.ES_CONTINUOUS | self.ES_SYSTEM_REQUIRED
            )
            self.is_preventing = True
            return True
        return False

    def allow(self):
        """恢复系统自动睡眠"""
        if self.is_preventing:
            ctypes.windll.kernel32.SetThreadExecutionState(self.ES_CONTINUOUS)
            self.is_preventing = False
            return True
        return False




# ==================== 倒计时弹窗 ====================
class AutoCloseMessageBox:
    """倒计时自动关闭的消息框"""
    
    def __init__(self, parent, title, message, auto_close_seconds=5, is_error=False):
        self.result = False
        self.toplevel = tk.Toplevel(parent)
        self.toplevel.title(title)
        self.toplevel.resizable(False, False)
        
        # ✅ 确保弹窗绝对置顶
        self.toplevel.attributes('-topmost', True)
        
        # 获取屏幕尺寸以居中
        self.toplevel.update_idletasks()
        width = int(380 * DPI_SCALE) # ✅ 也适配 DPI
        height = int(180 * DPI_SCALE) # ✅ 也适配 DPI
        x = (self.toplevel.winfo_screenwidth() - width) // 2
        y = (self.toplevel.winfo_screenheight() - height) // 2
        self.toplevel.geometry(f"{width}x{height}+{x}+{y}")
        
        icon = "❌" if is_error else "✅"
        tk.Label(self.toplevel, text=icon, font=("Segoe UI", 40)).pack(pady=5)
        
        tk.Label(self.toplevel, text=message, font=("微软雅黑", 10), 
                wraplength=320, justify="center").pack()
        
        self.countdown_label = tk.Label(
            self.toplevel, 
            text=f"自动关闭：{auto_close_seconds}秒", 
            font=("微软雅黑", 9), 
            fg="gray"
        )
        self.countdown_label.pack(pady=5)
        
        btn_frame = tk.Frame(self.toplevel)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="立即关闭", command=self._close, 
                 bg="#2196F3", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        
        # 在打包所有控件后，再次强制置顶并聚焦
        self.toplevel.update_idletasks()
        self.toplevel.lift()
        self.toplevel.focus_force()
        
        self.toplevel.transient(parent)
        self.toplevel.grab_set()
        
        # 启动倒计时
        self.remaining = auto_close_seconds
        self._countdown()
        
        parent.wait_window(self.toplevel)
    
    def _countdown(self):
        if self.remaining > 0:
            self.remaining -= 1
            self.countdown_label.config(text=f"自动关闭：{self.remaining}秒")
            self.toplevel.after(1000, self._countdown)
        else:
            self._close()
    
    def _close(self):
        try:
            self.toplevel.destroy()
        except:
            pass


# ==================== 定时任务管理窗口  ====================
class TaskSchedulerWindow:
    """
    简化的任务管理窗口：
    仅用于管理“开机/登录时自动启动主程序”的计划任务。
    """
    
    def __init__(self, parent, script_dir, power_manager=None):
        self.parent = parent
        self.script_dir = script_dir
        self.power_manager = power_manager
        
        self.window = tk.Toplevel(parent.root)
        self.window.title("开机自启设置")
        self.window.withdraw()
        
        # ✅ 修复闪烁：设置背景色为系统默认灰色，避免白色闪光
        self.window.configure(bg='#f0f0f0') 
        
        # ✅ 修复尺寸：根据 DPI 缩放
        win_width = int(400 * DPI_SCALE)
        win_height = int(300 * DPI_SCALE)
        
        # ✅ 关键：先设置几何位置，再显示，减少重绘
        self.window.geometry(f"{win_width}x{win_height}") 
        self.window.resizable(False, False)
        
        # ✅ 修复焦点：确保子窗口获取焦点，防止主窗口干扰
        self.window.transient(parent.root)
        self.window.grab_set()
        
        # ✅ 修复字体：应用全局 DPI 字体
        scaled_font_size = max(9, int(9 * DPI_SCALE))
        default_font = ("微软雅黑", scaled_font_size)
        self.window.option_add("*Font", default_font)
        self.window.option_add("*Background", '#f0f0f0') # 统一背景色

        # ... (后续代码保持不变: task_name, script_path 等初始化) ...
        self.task_name = SCHEDULED_TASK_NAME
        main_app_path = sys.argv[0]
        if not os.path.isabs(main_app_path):
            main_app_path = os.path.abspath(main_app_path)
        self.script_path = main_app_path

        self.enable_grace_checkin = tk.BooleanVar(value=bool(self.parent.config_manager.get("enable_grace_checkin", False)))
        self.grace_minutes_var = tk.StringVar(value=str(self.parent.config_manager.get("grace_period_minutes", 30)))
        self.last_valid_grace = self.grace_minutes_var.get()

        self.create_widgets()
        
        # ✅ 关键：所有控件加载完毕后，再显示窗口
        self.window.update_idletasks() # 强制计算布局
        self.window.deiconify()        # 显示窗口
        
        self.window.protocol("WM_DELETE_WINDOW", self._save_and_close)
        self.window.after(50, self.query_task_status)
    
    def _toggle_grace_input_state(self):
        """根据复选框状态，启用或禁用时间输入框"""
        if self.enable_grace_checkin.get():
            self.entry_grace_time.config(state="normal", bg="white")
        else:
            self.entry_grace_time.config(state="disabled", bg="#f0f0f0")

    def _save_and_close(self):
        """保存配置并关闭窗口"""
        try:
            # 验证输入
            minutes = int(self.grace_minutes_var.get())
            if minutes < 1 or minutes > 120:
                messagebox.showwarning("警告", "宽限时间必须在 1-120 分钟之间")
                return
            
            # 保存到 ConfigManager
            self.parent.config_manager.set("enable_grace_checkin", self.enable_grace_checkin.get())
            self.parent.config_manager.set("grace_period_minutes", minutes)
            self.parent.config_manager.save_config()
            
            self.log(f"[配置] 补打卡设置已保存: 启用={self.enable_grace_checkin.get()}, 宽限={minutes}分钟")
            self.window.destroy()
            
        except ValueError:
            messagebox.showwarning("警告", "请输入有效的数字")

    def create_widgets(self):
        info_text = (
            "此功能用于设置 Windows 登录时自动启动本程序。\n"
            "程序启动后将自动恢复上次主页设置的打卡状态。"
        )
        tk.Label(self.window, text=info_text, justify="left", font=("微软雅黑", 10),
                 wraplength=380).pack(pady=10, padx=10, anchor="w")
        
        status_frame = tk.Frame(self.window)
        status_frame.pack(fill="x", padx=10, pady=5)
        
        tk.Label(status_frame, text="当前状态:", font=("微软雅黑", 9, "bold")).pack(side="left")
        self.status_label = tk.Label(status_frame, text="检测中...", font=("微软雅黑", 9), fg="gray")
        self.status_label.pack(side="left", padx=5)

        # ✅ 新增：预创建路径警告容器和标签
        # 使用一个 Frame 包裹，方便整体 pack_forget
        self.warning_container = tk.Frame(self.window, bg="#ffebee")
        
        # 【修改点 1】计算动态换行长度，适配不同 DPI
        dynamic_wrap_length = int(380 * DPI_SCALE)

        # 【修改点 2】优化 Label 配置
        self.path_warning_lbl = tk.Label(
            self.warning_container,
            text="",
            font=("微软雅黑", 8),
            fg="#d32f2f",
            bg="#ffebee",
            wraplength=dynamic_wrap_length,  # 【修改】使用动态计算的换行长度
            justify="left",                  # 保持左对齐多行文本
            anchor="w",                      # 【新增】关键：让单行文本也靠左对齐
            padx=5,
            pady=5
        )
        # 【修改点 3】pack 时增加 fill="x" 和 expand=True
        # 这样 Label 会横向撑满 warning_container，文字会自动换行，不会挤占其他空间
        self.path_warning_lbl.pack(fill="x", expand=True) 
        
        # 初始不 pack warning_container，等到 query_task_status 中根据需要 pack

        # ✅ 补打卡设置区域
        self.grace_frame = tk.LabelFrame(self.window, text="错过打卡补救设置", padx=10, pady=10)
        # ... (后续代码保持不变)
        self.grace_frame.pack(fill="x", padx=10, pady=5)
        
        # 复选框
        self.chk_grace = tk.Checkbutton(
            self.grace_frame, 
            text="启用错过打卡自动补救", 
            variable=self.enable_grace_checkin,
            font=("微软雅黑", 9),
            command=self._toggle_grace_input_state # 绑定切换事件
        )
        self.chk_grace.pack(anchor="w")
        
        # 说明文字
        tk.Label(
            self.grace_frame, 
            text="若程序启动时已过设定打卡时间，且在宽限期内，将立即执行一次打卡。",
            font=("微软雅黑", 9), fg="#666", justify="left"
        ).pack(anchor="w", pady=(0, 5))
        
        # 时间输入行
        time_input_frame = tk.Frame(self.grace_frame)
        time_input_frame.pack(fill="x", pady=2)
        
        tk.Label(time_input_frame, text="宽限时间(分钟):", font=("微软雅黑", 9)).pack(side=tk.LEFT)
        
        self.entry_grace_time = tk.Spinbox(
            time_input_frame, 
            from_=1, to=120, width=5, 
            textvariable=self.grace_minutes_var,
            font=("微软雅黑", 9)
        )
        self.entry_grace_time.pack(side=tk.LEFT, padx=5)
        
        # 初始化输入框状态
        self._toggle_grace_input_state()

        btn_frame = tk.Frame(self.window)
        btn_frame.pack(pady=15)
        
        self.btn_create = tk.Button(btn_frame, text="启用开机自启", width=12, 
                                    command=self.create_or_update_task,
                                    bg="#4CAF50", fg="white", font=("微软雅黑", 9))
        self.btn_create.pack(side="left", padx=5)
        
        self.btn_delete = tk.Button(btn_frame, text="禁用开机自启", width=12, 
                                    command=self.delete_task,
                                    bg="#f44336", fg="white", font=("微软雅黑", 9))
        self.btn_delete.pack(side="left", padx=5)
        
        tk.Button(btn_frame, text="关闭", width=8, command=self._save_and_close,
                  font=("微软雅黑", 9)).pack(side="left", padx=5)

    def log(self, message):
        if hasattr(self.parent, '_log'):
            self.parent._log(f"[自启设置] {message}")

    def query_scheduled_task(self, task_name):
        """
        查询计划任务状态，并尝试解析其配置的执行路径。
        返回: (exists: bool, info: str, configured_cmd: str or None)
        """
        cmd = ["schtasks", "/query", "/tn", task_name, "/fo", "list", "/v"]
        try:
            if sys.platform == 'win32':
                CREATE_NO_WINDOW = 0x08000000
                result = subprocess.run(cmd, capture_output=True, text=True, check=False, creationflags=CREATE_NO_WINDOW, encoding='gbk')
            else:
                result = subprocess.run(cmd, capture_output=True, text=True, check=False)
                
            if result.returncode == 0:
                output = result.stdout
                configured_cmd = None
                
                # 解析 "Task To Run:" 或 "运行任务:"
                for line in output.splitlines():
                    # 兼容中英文 Windows
                    if "Task To Run:" in line or "运行任务:" in line or "要运行的任务:" in line:
                        # 分割一次，取后半部分
                        parts = line.split(":", 1)
                        if len(parts) > 1:
                            configured_cmd = parts[1].strip()
                        break
                        
                return True, output, configured_cmd
            else:
                return False, "任务不存在", None
        except Exception as e:
            return False, str(e), None

    def _check_path_consistency(self, configured_cmd):
        """
        检查配置的任务命令是否包含当前脚本路径。
        返回: (is_consistent: bool, warning_msg: str or None)
        """
        if not configured_cmd:
            return True, None
            
        # 当前脚本的绝对路径
        current_script_path = os.path.abspath(self.script_path)
        
        # 清理配置命令中的引号，方便比对
        # schtasks 返回的通常是: "C:\Python\pythonw.exe" "C:\Path\To\Script.py"
        # 我们只需要确认 current_script_path 存在于 configured_cmd 中即可
        
        if current_script_path in configured_cmd:
            return True, None
        else:
            msg = (
                "⚠️ 注意：检测到程序文件位置已移动！\n"
                "请先点击【禁用开机自启】，再点击【启用开机自启】以更新路径。"
            )
            return False, msg

    def query_task_status(self):
        """
        查询计划任务状态并更新UI，同时校验路径一致性。
        """
        exists, info, configured_cmd = self.query_scheduled_task(self.task_name)
        
        # 1. 先隐藏警告容器
        self.warning_container.pack_forget()

        if exists:
            # 状态：已启用
            self.status_label.config(text="已启用", fg="green")
            self.btn_create.config(state="disabled", text="已启用")
            self.btn_delete.config(state="normal")
            
            # ✅ 校验路径
            is_ok, warning_msg = self._check_path_consistency(configured_cmd)
            
            if not is_ok and warning_msg:
                # 更新警告文本
                self.path_warning_lbl.config(text=warning_msg)
                # 将警告容器插入到 status_frame 和 grace_frame 之间
                self.warning_container.pack(fill="x", padx=10, pady=(0, 5), before=self.grace_frame)
        else:
            # 状态：未配置
            self.status_label.config(text="未配置", fg="red")
            self.btn_create.config(state="normal", text="启用开机自启")
            self.btn_delete.config(state="disabled")

        # ✅ 新增：每次状态查询后，重新调整窗口大小以适应内容
        self._resize_window_to_fit_content()

    def _resize_window_to_fit_content(self):
        """
        根据当前可见控件的内容，动态调整窗口高度，避免下方留白或控件被遮挡。
        """
        try:
            # 1. 强制更新所有控件的几何信息，确保 winfo_reqheight() 准确
            self.window.update_idletasks()
            
            # 2. 计算所有 pack 过的控件的总高度
            # 注意：这里我们估算一个基础高度，或者更简单地，直接获取内容所需的总高度
            # 由于 pack 布局比较复杂，最稳妥的方法是获取最后一个控件的位置
            
            # 获取窗口中所有子控件
            widgets = self.window.winfo_children()
            
            max_y = 0
            for widget in widgets:
                # 只计算可见且已映射的控件
                if widget.winfo_ismapped():
                    # y位置 + 高度 + 可能的底部padding
                    h = widget.winfo_height()
                    y = widget.winfo_y()
                    # 获取 widget 的 pack 选项中的 pady (如果有)
                    info = widget.pack_info()
                    pady_top = int(info.get('pady', (0, 0))[0]) if isinstance(info.get('pady'), tuple) else int(info.get('pady', 0))
                    pady_bottom = int(info.get('pady', (0, 0))[1]) if isinstance(info.get('pady'), tuple) else int(info.get('pady', 0))
                    
                    current_bottom = y + h + pady_bottom
                    if current_bottom > max_y:
                        max_y = current_bottom

            # 3. 加上一些额外的底部边距，防止贴底
            new_height = max_y + 20 
            
            # 4. 限制最小和最大高度，防止极端情况
            min_h = int(320 * DPI_SCALE) # 最小高度
            max_h = int(450 * DPI_SCALE) # 最大高度（防止警告文字过多撑爆）
            
            if new_height < min_h:
                new_height = min_h
            elif new_height > max_h:
                new_height = max_h
                
            # 5. 保持宽度不变，只修改高度
            current_width = self.window.winfo_width()
            if current_width < 10: # 防止初始化时宽度为0
                current_width = int(400 * DPI_SCALE)
                
            # 6. 应用新尺寸
            self.window.geometry(f"{current_width}x{new_height}")
            
        except Exception as e:
            # 如果计算失败，忽略，不影响主功能
            pass

    def create_or_update_task(self):
        if not self.parent.power_manager.is_admin():
            messagebox.showwarning("权限不足", "需要管理员权限才能修改计划任务")
            return

        pythonw_path = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        if not os.path.exists(pythonw_path):
            pythonw_path = sys.executable
            
        cmd = ["schtasks", "/create", "/tn", self.task_name,
               "/tr", f'"{pythonw_path}" "{self.script_path}"',
               "/sc", "ONLOGON",
               "/f",
               "/rl", "HIGHEST",
               "/it"]
        
        self.log(f"正在创建登录自启任务...")
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, check=False, encoding='gbk')
            if result.returncode != 0:
                raise Exception(result.stderr)
            
            # ✅ 新增：使用 win32com 修改电源选项，取消“仅交流电”限制
            self._configure_task_power_settings()
            
            self.log("✅ 开机自启任务创建成功")
            messagebox.showinfo("成功", "开机自启已启用！\n下次登录时将自动启动本程序。")
            self.query_task_status()
            
        except Exception as e:
            self.log(f"❌ 创建失败: {e}")
            messagebox.showerror("失败", f"创建任务失败:\n{e}")

    def _configure_task_power_settings(self):
        """
        使用 COM 接口修改计划任务的电源设置：
        1. 允许在电池模式下运行
        2. 切换到电池模式时不停止
        """
        try:
            # 连接任务计划服务
            scheduler = win32com.client.Dispatch("Schedule.Service")
            scheduler.Connect()
            
            # 获取根文件夹
            root_folder = scheduler.GetFolder("\\")
            
            # 获取刚创建的任务
            task = root_folder.GetTask(self.task_name)
            task_def = task.Definition
            
            # 修改电源设置
            # DisallowStartIfOnBatteries = False  -> 允许在电池模式下启动
            task_def.Settings.DisallowStartIfOnBatteries = False
            
            # StopIfGoingOnBatteries = False      -> 切换到电池模式时不停止
            task_def.Settings.StopIfGoingOnBatteries = False
            
            # 保存更改 (6 = TASK_UPDATE)
            root_folder.RegisterTaskDefinition(
                self.task_name,
                task_def,
                6, 
                "", 
                "", 
                3  # TASK_LOGON_INTERACTIVE_TOKEN (保持原有的交互令牌登录类型)
            )
            self.log("[电源设置] 已配置为允许电池运行且切换电源时不中断")
            
        except Exception as e:
            self.log(f"[警告] 任务已创建，但高级电源设置修改失败: {e}")
            self.log("[提示] 任务仍可运行，但可能受系统默认电源策略限制")

    def delete_task(self):
        if not messagebox.askyesno("确认", "确定要禁用开机自启吗？\n下次登录后程序将不会自动启动。"):
            return

        if not self.parent.power_manager.is_admin():
            messagebox.showwarning("权限不足", "需要管理员权限")
            return

        cmd = ["schtasks", "/delete", "/tn", self.task_name, "/f"]
        try:
            # ✅ 修复：这里之前漏掉了 creationflags，导致删除任务时黑框闪烁
            if sys.platform == 'win32':
                CREATE_NO_WINDOW = 0x08000000
                result = subprocess.run(cmd, capture_output=True, text=True, check=False, creationflags=CREATE_NO_WINDOW)
            else:
                result = subprocess.run(cmd, capture_output=True, text=True, check=False)

            if result.returncode == 0:
                self.log("✅ 开机自启任务已删除")
                messagebox.showinfo("成功", "开机自启已禁用。")
                self.query_task_status()
            else:
                raise Exception(result.stderr)
        except Exception as e:
            self.log(f"❌ 删除失败: {e}")
            messagebox.showerror("失败", f"删除任务失败:\n{e}")
    


# ==================== 配置管理器 ====================
class ConfigManager:
    """配置管理器 - 保存和加载用户设置"""
    
    def __init__(self, config_file):
        self.config_file = config_file
        self.config = self._load_config()
    
    def _load_config(self):
        default_config = {
            "schedule_hour": 21,              # 默认打卡小时 (0-23)
            "schedule_minute": 00,            # 默认打卡分钟 (0-59)
            "window_geometry": f"{BASE_WIDTH}x{BASE_HEIGHT}", # 窗口大小 (通常不需改)
            "last_run": None,                 # 上次运行时间
            "custom_desktop_path": "",        # 自定义桌面路径
            "custom_keywords": [],
            "custom_mini_program_path": "",            # 自定义快捷方式关键词
            "enable_wechat_notify": False,    # 默认是否开启微信通知 (True/False)
            "close_wechat_after_notify": False, 
            "wechat_post_action": "无操作",   # 默认微信后续操作 ("无操作", "关闭窗口", "退出微信")
            "log_expanded": True,             # 默认日志区域是否展开
            "power_original_settings": None,  
            "power_settings_captured": False, 
            "is_timer_enabled": False,        # 默认定时任务状态 (True/False)
            "grace_period_minutes": 15,       # 默认错过打卡宽限时间 (分钟)
            "enable_grace_checkin": False     # 是否启用“错过打卡自动补救”功能
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    default_config.update(loaded)
                default_config["schedule_hour"] = int(default_config.get("schedule_hour", 21))
                default_config["schedule_minute"] = int(default_config.get("schedule_minute", 44))
        except Exception as e:
            print(f"[配置] 加载失败：{e}")
        
        return default_config
    
    def save_config(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"[配置] 保存失败：{e}")
            return False
    
    def get(self, key, default=None):
        return self.config.get(key, default)
    
    def set(self, key, value):
        self.config[key] = value
    
    def save_and_get(self, key, value):
        self.config[key] = value
        return self.save_config()
    
    def save_power_original_settings(self, settings):
        self.config["power_original_settings"] = settings
        self.config["power_settings_captured"] = True
        return self.save_config()
    
    def get_power_original_settings(self):
        if self.config.get("power_settings_captured", False):
            return self.config.get("power_original_settings")
        return None
    
    def reset_power_original_settings(self):
        self.config["power_original_settings"] = None
        self.config["power_settings_captured"] = False
        return self.save_config()

# ==================== 电源管理器 (极简版) ====================
class PowerManager:
    """
    由于采用了 ApiSleepPreventer，此类不再负责修改电源 GUID。
    仅保留 is_admin 工具方法供计划任务窗口使用。
    """
    def __init__(self, logger=None, config_manager=None):
        self._logger = logger

    def _log(self, msg):
        if self._logger:
            self._logger(msg)

    @staticmethod
    def is_admin():
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def save_and_disable_sleep(self, *args, **kwargs):
        # 废弃方法，不再执行任何操作
        return True

    def restore_settings(self, *args, **kwargs):
        # 废弃方法，不再执行任何操作
        return True


# ==================== 主界面 ====================
class AutoCheckInGUI:
    """自动打卡 GUI 管理程序"""
    
    # 全局唯一标识符
    MUTEX_NAME = "Global\\AutoCheckIn_System_Mutex_2025" 
    # 窗口类名通常由 Tkinter 自动生成，但我们可以通过标题查找
    WINDOW_TITLE = "某安自动打卡系统1.1 by ᥬ💯ᩤ"

    def __init__(self):
        self.time_config_changed = False 
        # --- 单例检测 ---
        self.h_mutex = None
        try:
            self.h_mutex = win32event.CreateMutex(None, False, self.MUTEX_NAME)
            if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
                ctypes.windll.user32.MessageBoxW(
                    0, 
                    "程序已在后台运行！\n\n请检查系统托盘图标（右下角）。", 
                    "提示", 
                    0x40 | 0x0
                )
                os._exit(0) 
        except Exception as e:
            print(f"[错误] 单例检测异常: {e}")
        # --------------

        self.root = tk.Tk()
        # ✅ 新增：根据 DPI 设置默认字体大小，避免文字过小
        # 计算缩放后的字体大小 (基础 7pt * 缩放比例)
        scaled_font_size = max(7, int(7 * DPI_SCALE))
        default_font = ("微软雅黑", scaled_font_size)
        
        # 设置全局默认字体
        self.root.option_add("*Font", default_font)
        # 设置 LabelFrame 标题字体稍大一点
        self.root.option_add("*LabelFrame*Font", ("微软雅黑", scaled_font_size + 1, "bold"))
        self.root.title("某安自动打卡系统1.1 by ᥬ💯ᩤ")
        self.root.geometry(f"{BASE_WIDTH}x{BASE_HEIGHT}")
        self.root.resizable(True, True)
        
        self.config_manager = ConfigManager(CONFIG_FILE)
        self.shortcut_detected = False
        
        # ✅ 修复1：补全缺失的变量初始化
        self.last_timer_state = bool(self.config_manager.get("is_timer_enabled", False))
       
        self.wechat_action_var = tk.StringVar(value="无操作")
        self.log_expanded = tk.BooleanVar(value=bool(self.config_manager.get("log_expanded", True)))

        self.schedule_hour = tk.StringVar(value=f"{int(self.config_manager.get('schedule_hour', 21)):02d}")
        self.schedule_minute = tk.StringVar(value=f"{int(self.config_manager.get('schedule_minute', 44)):02d}")
        self.last_valid_hour = self.schedule_hour.get()
        self.last_valid_minute = self.schedule_minute.get()
        
        self.enable_wechat_notify = tk.BooleanVar(value=bool(self.config_manager.get("enable_wechat_notify", False)))
 
        self.expanded_geometry = f"{BASE_WIDTH}x{BASE_HEIGHT}"
        self.timer_thread = None
        self.is_timer_running = False
        self.stop_timer_flag = False
        
        self.current_process = None
        self.process_lock = threading.Lock()
        self.is_checkin_running = False
        self.is_closing = False
        self.skip_next_grace_check = False
        self.last_aborted_date = None
        self.today_checkin_attempted = None 
        self.api_sleep_preventer = ApiSleepPreventer()
        self.power_manager = PowerManager(logger=self._log, config_manager=self.config_manager)
        self.is_scheduled_task = "--scheduled-task" in sys.argv
        self.scheduled_task_name = SCHEDULED_TASK_NAME
        
        self.root.after(500, self._check_power_state_on_startup)
        
        self._create_widgets()
        self.root.after(1500, self._restore_timer_state_on_startup)
        self.root.after(100, lambda: self._check_shortcut_on_startup())

    def _calculate_next_target_and_check_missed(self, hour, minute):
        """
        计算下一个目标时间，并检查是否刚刚错过（在宽限期内）。
        返回: (should_checkin_now: bool, next_target_time: datetime)
        """
        now = datetime.now()
        today_target = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        
        # 获取宽限期配置，默认30分钟
        grace_minutes = self.config_manager.get("grace_period_minutes", 30)
        
        # 情况1: 当前时间还没到今天的打卡时间
        if now < today_target:
            return False, today_target
            
        # 情况2: 当前时间已经过了今天的打卡时间
        # 计算过了多久
        diff_seconds = (now - today_target).total_seconds()
        grace_seconds = grace_minutes * 60
        
        if diff_seconds <= grace_seconds:
            # 在宽限期内，应该立即补打卡
            # 下一次目标时间是明天
            tomorrow_target = today_target + timedelta(days=1)
            return True, tomorrow_target
        else:
            # 超过宽限期，今天放弃，等待明天
            tomorrow_target = today_target + timedelta(days=1)
            return False, tomorrow_target


    def _create_tray_icon(self):
        """创建系统托盘图标"""
        icon_path = os.path.join(SCRIPT_DIR, "ciga.ico")
        image = None
        
        try:
            if os.path.exists(icon_path):
                img = Image.open(icon_path)
                if img.mode != 'RGBA':
                    img = img.convert('RGBA')
                img = img.resize((64, 64), Image.LANCZOS)
                image = img
                self._log(f"[托盘] 已加载自定义图标: {icon_path}")
            else:
                raise FileNotFoundError("未找到ciga.ico")
        except Exception as e:
            self._log(f"[托盘] 加载图标失败 ({e})，使用默认图标")
            image = Image.new('RGBA', (64, 64), color=(0, 120, 215, 255))
            draw = ImageDraw.Draw(image)
            draw.ellipse((10, 10, 54, 54), fill='white')

        menu = pystray.Menu(
            pystray.MenuItem('显示主界面', self._show_main_window, default=True),
            pystray.MenuItem('立即打卡', self._tray_immediate_checkin),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('退出程序', self._on_closing_from_tray)
        )
        
        self.tray_icon = pystray.Icon("AutoCheckIn", image, "某安自动打卡系统1.1", menu, on_double_click_left=self._show_main_window)
        
        if not hasattr(self, 'tray_thread') or not self.tray_thread.is_alive():
            self.tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
            self.tray_thread.start()
            self._log("[托盘] 系统托盘已启动")

    def _show_main_window(self, icon=None, item=None):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        if sys.platform == 'win32':
            self.root.attributes('-topmost', True)
            self.root.after(100, lambda: self.root.attributes('-topmost', False))

    def _hide_to_tray(self):
        self.root.withdraw()
        self._log("[托盘] 主窗口已隐藏到托盘")

    def _tray_immediate_checkin(self, icon=None, item=None):
        self.root.after(0, self._immediate_checkin)

    def _on_closing_from_tray(self, icon=None, item=None):
        self.is_closing = True 
        self.root.after(0, self._on_closing)


    def _create_widgets(self):
        # ===== 标题区域 =====
        title_frame = tk.Frame(self.root)
        title_frame.pack(pady=2)
        tk.Label(title_frame, text="某安自动打卡系统", font=("微软雅黑", 14, "bold")).pack()
        
        # ===== 第一行：微信通知 | 系统状态 =====
        row1_frame = tk.Frame(self.root)
        row1_frame.pack(fill="x", padx=5, pady=2)
        row1_frame.grid_columnconfigure(0, weight=2)
        row1_frame.grid_columnconfigure(1, weight=2)
        
        notify_frame = tk.LabelFrame(row1_frame, text="微信通知", padx=5, pady=5)
        notify_frame.grid(row=0, column=0, sticky="nsew", padx=2)
        tk.Checkbutton(notify_frame, text="微信通知 - 发送打卡情况（可能不适配）", variable=self.enable_wechat_notify,
                    font=("微软雅黑", 9), command=self._save_notify_settings).pack(anchor="w")
        
        close_wechat_frame = tk.Frame(notify_frame)
        close_wechat_frame.pack(fill="x", pady=2)
        tk.Label(close_wechat_frame, text="打卡后对微信:", font=("微软雅黑", 9)).pack(side=tk.LEFT)
        
                # ✅ 新增：创建下拉框
        self.combo_wechat_action = ttk.Combobox(
            close_wechat_frame, 
            textvariable=self.wechat_action_var,
            values=["无操作", "关闭窗口", "退出微信"],
            state="readonly",
            width=10,
            font=("微软雅黑", 9)
        )
        self.combo_wechat_action.pack(side=tk.LEFT, padx=5)
        # 绑定事件：当选择改变时，自动保存配置
        self.combo_wechat_action.bind("<<ComboboxSelected>>", self._save_wechat_action_config)
        status_frame = tk.LabelFrame(row1_frame, text="系统状态", padx=5, pady=5)
        status_frame.grid(row=0, column=1, sticky="nsew", padx=2)
        self.status_label = tk.Label(status_frame, text="状态：待机", font=("微软雅黑", 8))
        self.status_label.pack(anchor="w")
        self.next_time_label = tk.Label(status_frame, text="下次打卡：未设置", font=("微软雅黑", 8))
        self.next_time_label.pack(anchor="w")
        self.power_status_label = tk.Label(status_frame, text="电源：检测中...", font=("微软雅黑", 8), fg="gray")
        self.power_status_label.pack(anchor="w")
        
        # ===== 第二行：重要提示 | 打卡时间 =====
        row2_frame = tk.Frame(self.root)
        row2_frame.pack(fill="x", padx=5, pady=2)
        row2_frame.grid_columnconfigure(0, weight=2)
        row2_frame.grid_columnconfigure(1, weight=2)
        
        notice_frame = tk.LabelFrame(row2_frame, text="重要提示", padx=5, pady=5)
        notice_frame.grid(row=0, column=0, sticky="nsew", padx=2)
        notices = [
            "1 保持微信登录  2 添加小程序桌面快捷方式",
            "3 长期定时打卡建议'连接电源'"
        ]
        for text in notices:
            tk.Label(notice_frame, text=text, font=("微软雅黑", 8)).pack(anchor="w")
        
        time_frame = tk.LabelFrame(row2_frame, text="打卡时间（不要退出程序托盘）", padx=5, pady=5)
        time_frame.grid(row=0, column=1, sticky="nsew", padx=2)
        time_inner = tk.Frame(time_frame)
        time_inner.pack()
        tk.Label(time_inner, text="时:", font=("微软雅黑", 8)).grid(row=0, column=0)
        tk.Spinbox(time_inner, from_=0, to=23, width=3, textvariable=self.schedule_hour,
                format="%02.0f").grid(row=0, column=1, padx=2)
        tk.Label(time_inner, text="分:", font=("微软雅黑", 8)).grid(row=0, column=2)
        tk.Spinbox(time_inner, from_=0, to=59, width=3, textvariable=self.schedule_minute,
                format="%02.0f").grid(row=0, column=3, padx=2)
        tk.Button(time_frame, text="保存", command=self._save_time_settings,
                font=("微软雅黑", 8), bg="#28A745", fg="white").pack(pady=2)
        
        # ===== 第三行：登录设置 | 快捷方式检测 =====
        row3_frame = tk.Frame(self.root)
        row3_frame.pack(fill="x", padx=5, pady=2)
        row3_frame.grid_columnconfigure(0, weight=2)
        row3_frame.grid_columnconfigure(1, weight=2)
        
        login_frame = tk.LabelFrame(row3_frame, text="登录设置", padx=5, pady=5)
        login_frame.grid(row=0, column=0, sticky="nsew", padx=2)
        tk.Label(login_frame, text="关闭锁屏密码，否则'定时/计划'打卡一定失败",
                font=("微软雅黑", 8), fg="red").pack(anchor="w")
        self.btn_login_settings = tk.Button(login_frame, text="打开登录选项", command=self._open_login_settings,
                                            font=("微软雅黑", 8), bg="#2196F3", fg="white")
        self.btn_login_settings.pack(pady=2)
        
        shortcut_frame = tk.LabelFrame(row3_frame, text="桌面快捷方式", padx=5, pady=5)
        shortcut_frame.grid(row=0, column=1, sticky="nsew", padx=2)
        self.shortcut_status_label = tk.Label(shortcut_frame, text="检测状态：未检测", font=("微软雅黑", 8), fg="gray")
        self.shortcut_status_label.pack(anchor="w")
        # --- 新增：按钮容器，让两个按钮并排 ---
        btn_row_frame = tk.Frame(shortcut_frame)
        btn_row_frame.pack(pady=2)

        self.btn_check_shortcut = tk.Button(btn_row_frame, text="自动检测", command=lambda: self._check_shortcut(show_popup=True),
                                            font=("微软雅黑", 8), bg="#17A2B8", fg="white", width=8)
        self.btn_check_shortcut.pack(side=tk.LEFT, padx=2)

        # ✅ 新增：手动选择路径按钮
        self.btn_select_path = tk.Button(btn_row_frame, text="选择路径", command=self._select_mini_program_path,
                                         font=("微软雅黑", 8), bg="#FF9800", fg="white", width=8)
        self.btn_select_path.pack(side=tk.LEFT, padx=2)

        # ===== 操作按钮区域 =====
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=3)
        self.btn_immediate = tk.Button(btn_frame, text="立即打卡", command=self._immediate_checkin,
                                    width=10, bg="#4CAF50", fg="white")
        self.btn_immediate.pack(side=tk.LEFT, padx=5)
        self.btn_timer_toggle = tk.Button(btn_frame, text="定时打卡", command=self._toggle_timer,
                                  width=10, bg="#2196F3", fg="white")
        self.btn_timer_toggle.pack(side=tk.LEFT, padx=5)
       
        self.btn_scheduler = tk.Button(btn_frame, text="创建计划任务", command=self._open_scheduler_window,
                                       width=10)
        self.btn_scheduler.pack(side=tk.LEFT, padx=5)
        
        # ===== 日志显示区域 =====
        log_frame = tk.LabelFrame(self.root, text="运行日志", padx=5, pady=5)
        log_frame.pack(fill="both", expand=True, padx=5, pady=2)
        
        log_control_frame = tk.Frame(log_frame)
        log_control_frame.pack(fill="x", pady=2)
        
        self.btn_toggle_log = tk.Button(log_control_frame, text="▼ 折叠日志", 
                                        command=self._toggle_log,
                                        font=("微软雅黑", 8), width=10,
                                        bg="#6C757D", fg="white")
        self.btn_toggle_log.pack(side=tk.LEFT)
        
        tk.Button(log_control_frame, text="清空日志", command=self._clear_log, 
                font=("微软雅黑", 8)).pack(side=tk.LEFT, padx=5)
        
        tk.Button(log_control_frame, text="关闭窗口", command=self._on_closing, 
                font=("微软雅黑", 8), bg="#dc3545", fg="white").pack(side=tk.RIGHT, padx=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, width=70, font=("Consolas", 8))
        
        if self.log_expanded.get():
            self.log_text.pack(fill="both", expand=True)
            self.btn_toggle_log.config(text="▼ 折叠日志")
            self.root.geometry(f"{BASE_WIDTH}x{BASE_HEIGHT}")
        else:
            self.btn_toggle_log.config(text="▶ 展开日志")
            fold_height = int(BASE_HEIGHT * FOLD_RATIO)
            self.root.geometry(f"{BASE_WIDTH}x{fold_height}")
        
        self.root.after(100, lambda: self._check_shortcut_on_startup())
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _select_mini_program_path(self):
        """
        弹出文件选择框，让用户手动选择小程序的 .lnk 快捷方式
        """
        file_path = filedialog.askopenfilename(
            title="请选择小程序快捷方式 (.lnk)",
            initialdir=os.path.expanduser("~\\Desktop"), # 默认打开桌面
            filetypes=[("快捷方式", "*.lnk"), ("所有文件", "*.*")]
        )
        
        if file_path:
            # 1. 保存到配置文件
            self.config_manager.set("custom_mini_program_path", file_path)
            self.config_manager.save_config()
            
            # 2. 更新 UI 状态
            self.shortcut_status_label.config(text=f"已指定: {os.path.basename(file_path)}", fg="green", font=("微软雅黑", 8, "bold"))
            self._log(f"[路径设置] 用户手动指定小程序路径: {file_path}")
            
            # 3. 标记为已检测到，避免启动时报错
            self.shortcut_detected = True
            
            messagebox.showinfo("成功", "路径已保存！\n程序将优先使用此路径启动小程序。")
    
    def _open_scheduler_window(self):
        try:
            if not self.power_manager.is_admin():
                messagebox.showwarning("权限不足", 
                    "创建/删除计划任务 任务需要管理员权限！\n\n"
                    "请右键点击程序 -> 以管理员身份运行")
                return
            
            scheduler_window = TaskSchedulerWindow(self, SCRIPT_DIR, self.power_manager)
            self._log("[计划任务] 已打开定时计划管理窗口")
        except Exception as e:
            self._log(f"[计划任务] 打开窗口失败：{e}")
            messagebox.showerror("错误", f"无法打开计划任务窗口:\n{str(e)}")

    def _check_power_state_on_startup(self):
        try:
            self._update_power_status()
        except Exception as e:
            self._log(f"[启动检测] 异常：{e}")
            self._update_power_status()

    def _has_scheduled_task(self):
        try:
            cmd = ["schtasks", "/query", "/tn", self.scheduled_task_name, "/fo", "list"]
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
            exists = result.returncode == 0
            return exists
        except Exception as e:
            return False
    
    def _update_power_status(self):
        is_timer = self.is_timer_running
        if is_timer:
            self.power_status_label.config(text="电源：防睡眠中", fg="red")
        else:
            self.power_status_label.config(text="电源：正常（允许睡眠）", fg="green")
        self.root.update_idletasks()

    
    def _restore_timer_state_on_startup(self):
        if self.is_closing:
            return

        self._log("[启动恢复] 检查定时任务状态...")
        
        if self.last_timer_state:
            self._log("[启动恢复] 检测到上次定时任务为【开启】状态，正在自动恢复...")
            
            self.btn_timer_toggle.config(text="停止定时", bg="#ff4848", fg="white")
            self.is_timer_running = True 
            
            if self.api_sleep_preventer.prevent():
                self._log("[电源] ✅ API 防睡眠已自动启用")
            else:
                self._log("[电源] ⚠️ API 防睡眠启用失败")
            
            hour = int(self.schedule_hour.get())
            minute = int(self.schedule_minute.get())
            
            # ✅ 使用新逻辑计算目标和是否补打卡
            should_checkin_now, next_target = self._calculate_next_target_and_check_missed(hour, minute)
            
            # ✅ 新增：检查是否启用了补打卡功能
            enable_grace = self.config_manager.get("enable_grace_checkin", False)
            
            if should_checkin_now and enable_grace:
                # ✅ 关键修改：标记正在运行，并更新按钮状态
                self.is_checkin_running = True
                self._toggle_immediate_button(to_terminate_mode=True)
                
                # 立即在后台线程执行补打卡
                self._update_status("正在补打卡...", next_target.strftime("%Y-%m-%d %H:%M:%S"))
                thread = threading.Thread(target=self._run_checkin, args=(True, False), daemon=True)
                thread.start()
            elif should_checkin_now and not enable_grace:
                 self._log("[启动恢复] 检测到错过打卡，但‘自动补救’未启用，跳过本次补打")
                 self._update_status("定时运行中", next_target.strftime("%Y-%m-%d %H:%M:%S"))
            else:
                # 正常等待
                self._update_status("定时运行中", next_target.strftime("%Y-%m-%d %H:%M:%S"))
                self._log(f"[启动恢复] 定时任务已恢复，下次打卡：{next_target.strftime('%Y-%m-%d %H:%M:%S')}")
            
            # 启动主定时循环线程
            self.timer_thread = threading.Thread(target=self._timer_loop, args=(hour, minute), daemon=True)
            self.timer_thread.start()
            
            self._update_power_status()
        else:
            self._log("[启动恢复] 上次定时任务为【关闭】状态，保持待机")
            self._update_power_status()

    def _check_shortcut_on_startup(self):
        # ✅ 新增：优先检查是否有用户自定义路径
        custom_path = self.config_manager.get("custom_mini_program_path", "")
        if custom_path and os.path.exists(custom_path):
            self.shortcut_detected = True
            self.shortcut_status_label.config(text=f"已指定: {os.path.basename(custom_path)}", fg="green", font=("微软雅黑", 8, "bold"))
            self._log(f"[启动] 加载用户自定义小程序路径: {custom_path}")
            return # 如果有自定义路径且存在，直接返回，不再执行下面的自动检测

        # 原有的自动检测逻辑
        self._check_shortcut(show_popup=False)
        if not self.shortcut_detected and not self.is_closing:
            self._log("[快捷方式] 初次启动未检测到桌面快捷方式")
            self._show_shortcut_warning(allow_continue=False)
        
        
        saved_action = self.config_manager.get("wechat_post_action", "无操作")
        self.wechat_action_var.set(saved_action)
        
        self._log("系统初始化完成")
        self._log(f"脚本目录：{SCRIPT_DIR}")
        hour = int(self.schedule_hour.get())
        minute = int(self.schedule_minute.get())
        self._log(f"配置已加载：{hour:02d}:{minute:02d}")
        self._log(f"微信通知：{'开启' if self.enable_wechat_notify.get() else '关闭'}")
        
        if not self.power_manager.is_admin():
            self._log("[警告] 未以管理员身份运行，部分功能可能受限")
    
    def _save_time_settings(self, show_message=True):
        try:
            hour = int(self.schedule_hour.get())
            minute = int(self.schedule_minute.get())
        except ValueError:
            if show_message:
                messagebox.showwarning("警告", "时间格式无效")
            return False
        
        if not (0 <= hour <= 23):
            if show_message:
                messagebox.showwarning("警告", "小时必须在 0-23 之间")
            self.schedule_hour.set(self.last_valid_hour)
            return False
        
        if not (0 <= minute <= 59):
            if show_message:
                messagebox.showwarning("警告", "分钟必须在 0-59 之间")
            self.schedule_minute.set(self.last_valid_minute)
            return False
        
        if self.config_manager.save_and_get("schedule_hour", hour) and \
        self.config_manager.save_and_get("schedule_minute", minute):
            self.last_valid_hour = f"{hour:02d}"
            self.last_valid_minute = f"{minute:02d}"
            self.schedule_hour.set(self.last_valid_hour)
            self.schedule_minute.set(self.last_valid_minute)
            self._log(f"[配置] 时间设置已保存：{hour:02d}:{minute:02d}")
            
            # ✅ 新增：如果定时正在运行，标记配置已更改，线程会在下一轮循环检测到
            if self.is_timer_running:
                self.time_config_changed = True
                self._log("[提示] 定时运行中，新时间将在下一个检查周期生效")
                # 可选：强制唤醒线程（如果线程在长睡眠中），这里简单处理，依靠短睡眠轮询即可
            
            if show_message:
                messagebox.showinfo("成功", f"时间设置已保存")
            return True
        else:
            return False
    
    
    def _save_wechat_action_config(self, event=None):
        """保存微信后续操作设置"""
        action = self.wechat_action_var.get()
        self.config_manager.save_and_get("wechat_post_action", action)
        # 可选：打印日志确认
        # self._log(f"[配置] 微信后续操作已设置为: {action}")

    def _save_notify_settings(self):
        value = self.enable_wechat_notify.get()
        self.config_manager.save_and_get("enable_wechat_notify", value)

    def _toggle_timer(self):
        if self.is_timer_running:
            self._stop_timer()
        else:
            self._start_timer()
   

    def _check_shortcut(self, show_popup=True):
        if self.is_closing:
            return
        
        shortcut_found = False
        target_shortcut = None
        desktop_paths = []
        
        try:
            # ✅ 增强1：使用 shell32 获取真实的桌面路径，兼容 OneDrive 和重定向
            try:
                import winreg
                # 尝试从注册表获取 User Shell Folders 中的 Desktop 路径
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
                real_desktop, _ = winreg.QueryValueEx(key, "Desktop")
                winreg.CloseKey(key)
                # 展开环境变量 (如 %USERPROFILE%\Desktop)
                real_desktop = os.path.expandvars(real_desktop)
                if os.path.exists(real_desktop):
                    desktop_paths.append(real_desktop)
            except Exception:
                pass

            # ✅ 增强2：保留原有的常规路径作为备选
            user_profile = os.environ.get('USERPROFILE', '')
            if user_profile:
                std_desktop = os.path.join(user_profile, 'Desktop')
                if os.path.exists(std_desktop) and std_desktop not in desktop_paths:
                    desktop_paths.append(std_desktop)
            
            onedrive = os.environ.get('OneDrive', '')
            if onedrive:
                od_desktop = os.path.join(onedrive, 'Desktop')
                if os.path.exists(od_desktop) and od_desktop not in desktop_paths:
                    desktop_paths.append(od_desktop)

            # 去重
            desktop_paths = list(set(desktop_paths))
            
            if not desktop_paths:
                self._log("[检测] 未找到有效的桌面路径")
                self.shortcut_status_label.config(text="检测失败: 无桌面路径", fg="red")
                return

            self._log(f"[检测] 正在扫描以下路径: {desktop_paths}")

            custom_keywords = self.config_manager.get("custom_keywords", [])
            default_keywords = ['中南林业科技大学学生工作部']
            keywords = list(set(default_keywords + custom_keywords))
            
            for desktop in desktop_paths:
                if not os.path.exists(desktop):
                    continue
                try:
                    files = os.listdir(desktop)
                    for file in files:
                        if file.lower().endswith('.lnk'):
                            # 检查文件名是否包含关键词
                            for keyword in keywords:
                                if keyword.lower() in file.lower():
                                    shortcut_found = True
                                    # 优先标记完全匹配的
                                    if keyword == '中南林业科技大学学生工作部':
                                        target_shortcut = file
                                    break
                except PermissionError:
                    self._log(f"[检测] 权限不足，无法访问: {desktop}")
                except Exception as e:
                    self._log(f"[检测] 扫描 {desktop} 出错: {e}")

            self.shortcut_detected = shortcut_found
            
            # ✅ 增强3：明确更新 UI
            if shortcut_found:
                if target_shortcut:
                    msg = f"小程序: {target_shortcut}"
                    self.shortcut_status_label.config(text=msg, fg="green", font=("微软雅黑", 8, "bold"))
                    self._log(f"[检测] 成功: {msg}")
                else:
                    self.shortcut_status_label.config(text="已检测到相关快捷方式", fg="green")
                    self._log("[检测] 成功: 找到相关快捷方式")
            else:
                self.shortcut_status_label.config(text="未检测到相关快捷方式", fg="red")
                self._log("[检测] 失败: 未找到包含关键词的 .lnk 文件")
                
        except Exception as e:
            self._log(f"[检测] 发生未知错误: {e}")
            self.shortcut_status_label.config(text="检测异常", fg="red")
        
        # 如果是在点击按钮时触发，且未找到，可以选择是否弹窗
        if show_popup and not shortcut_found:
             # 可选：是否每次点击都弹窗？建议只在启动时弹窗，点击时只更新标签
             pass 
    
    def _show_shortcut_warning(self, allow_continue=True):
        warning_message = (
            "⚠️ 未检测到小程序桌面快捷方式\n\n"
            "请按以下步骤操作：\n"
            "1. 打开微信 -> 进入平安打卡小程序\n"
            "2. 右上角三个点创建桌面快捷方式\n"
            "3. 重新检测\n\n"
        )
        if allow_continue:
            return messagebox.askyesno("警告", warning_message, icon="warning")
        else:
            messagebox.showwarning("重要提示", warning_message, icon="warning")
            return False
    
    def _open_login_settings(self):
        try:
            subprocess.Popen(["start", "ms-settings:signinoptions"], shell=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开登录选项:\n{str(e)}")
    
    def _log(self, message):
        if not message or not message.strip():
            return
        timestamp = datetime.now().strftime("%H:%M:%S") 
        log_line = f"[{timestamp}] {message}\n"
        try:
            if self.is_closing or not self.log_text.winfo_exists():
                return
            self.log_text.insert(tk.END, log_line)
            self.log_text.see(tk.END)
        except Exception:
            pass

        # ✅ 新增：线程安全的日志转发方法
    def _log_to_gui_safe(self, message):
        """
        供打卡模块调用的线程安全日志接口。
        它不直接操作 UI，而是通过 root.after 将任务交给主线程执行。
        """
        # root.after(0, func) 是 Tkinter中跨线程更新UI的标准做法
        # 它会将 func 放入主事件队列，由主线程在下一个周期执行
        self.root.after(0, lambda: self._log(message))    
    def _toggle_log(self):
        if self.log_expanded.get():
            self.expanded_geometry = self.root.geometry().split('+')[0]
            self.log_text.pack_forget()
            self.btn_toggle_log.config(text="▶ 展开日志")
            self.log_expanded.set(False)
            fold_height = int(BASE_HEIGHT * FOLD_RATIO)
            self.root.geometry(f"{BASE_WIDTH}x{fold_height}")
            self._save_log_expanded_state()
        else:
            self.log_text.pack(fill="both", expand=True)
            self.btn_toggle_log.config(text="▼ 折叠日志")
            self.log_expanded.set(True)
            self.root.geometry(self.expanded_geometry)
            self._save_log_expanded_state()
    
    def _clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def _save_log_expanded_state(self):
        try:
            self.config_manager.set("log_expanded", self.log_expanded.get())
            self.config_manager.save_config()
        except Exception as e:
            pass
    
    def _save_all_config(self):
        try:
            self.config_manager.set("schedule_hour", self.schedule_hour.get())
            self.config_manager.set("schedule_minute", self.schedule_minute.get())
            self.config_manager.set("enable_wechat_notify", self.enable_wechat_notify.get())
            self.config_manager.set("log_expanded", self.log_expanded.get())
            self.config_manager.set("window_geometry", self.root.geometry())
            self.config_manager.save_config()
        except Exception as e:
            pass
    
    def _update_status(self, status, next_time=None):
        self.status_label.config(text=f"状态：{status}")
        if next_time:
            self.next_time_label.config(text=f"下次打卡：{next_time}")
        self.root.update_idletasks()
    

    
    def _toggle_immediate_button(self, to_terminate_mode=False):
        # ✅ 增加安全检查：如果当前有进程在运行，强制显示为终止模式
        with self.process_lock:
            is_process_alive = self.current_process is not None and self.current_process.poll() is None
        
        if to_terminate_mode or is_process_alive:
            self.btn_immediate.config(text="终止打卡", command=self._terminate_checkin,
                                     bg="#f44336", fg="white")
        else:
            self.btn_immediate.config(text="立即打卡", command=self._immediate_checkin,
                                     bg="#4CAF50", fg="white")



    def _immediate_checkin(self):
        """
        处理“立即打卡”按钮的点击事件。
        """
        # 1. 检查程序是否正在关闭
        if self.is_closing:
            self._log("[立即打卡] 程序正在关闭，忽略操作")
            return

        # 2. 检查定时任务状态
        if self.is_timer_running:
            self._log("[立即打卡] 定时任务正在运行，请先停止")
            messagebox.showwarning("警告", "定时任务正在运行，请先停止定时任务")
            return
        
        self._log("[立即打卡] 请求已接收，开始预处理...")

        # ✅ 关键修复：重置停止标志位，防止因之前停止定时导致的拦截
        if self.stop_timer_flag:
            self._log("[立即打卡] 检测到 stop_timer_flag 为 True，强制重置为 False")
            self.stop_timer_flag = False

        # 3. 检查快捷方式 (这是最常见的静默失败点)
        if not self.shortcut_detected:
            self._log("[立即打卡] 未检测到快捷方式，弹出警告...")
            # 尝试刷新一下检测状态，以防之前检测失败是由于时序问题
            self._check_shortcut(show_popup=False)
            
            if not self.shortcut_detected:
                user_choice = self._show_shortcut_warning(allow_continue=True)
                if not user_choice:
                    self._log("[立即打卡] 用户取消打卡（未确认快捷方式警告）")
                    return
                else:
                    self._log("[立即打卡] 用户强制继续打卡")
                    # ✅ 重要：如果用户强制继续，我们暂时认为快捷方式OK，避免后续重复弹窗
                    # 注意：这里不改变 self.shortcut_detected，以免干扰正常检测逻辑，
                    # 但既然通过了这里，就说明用户允许执行。
            else:
                self._log("[立即打卡] 重新检测发现快捷方式存在")

        # 4. 设置运行状态
        self.is_checkin_running = True
        self._toggle_immediate_button(to_terminate_mode=True)
        self._update_status("正在打卡...")
        self._log("[立即打卡] 启动后台打卡线程...")
        
        # 5. 在后台线程执行打卡，避免阻塞 UI
        try:
            thread = threading.Thread(target=self._run_checkin, args=(False, False), daemon=True)
            thread.start()
        except Exception as e:
            self._log(f"[立即打卡] 线程启动失败: {e}")
            self.is_checkin_running = False
            self._toggle_immediate_button(to_terminate_mode=False)
            messagebox.showerror("错误", f"无法启动打卡线程:\n{e}")

# 在 启动窗口二.py 的 AutoCheckInGUI 类中修改以下方法

    def _terminate_current_process(self):
        """
        强制终止当前可能卡住的进程。
        由于我们是 import 模式，没有 self.current_process 句柄，
        所以这里直接查杀常见的卡顿源：微信、小程序容器。
        """
        self._log("[终止] 正在清理可能卡住的进程...")
        
        # 需要查杀的进程列表
        targets = ["WeChat.exe", "Weixin.exe"] 
        
        for proc_name in targets:
            try:
                # 使用 taskkill /F /IM 强制杀死所有同名进程
                # CREATE_NO_WINDOW (0x08000000) 防止黑框闪烁
                subprocess.run(
                    ["taskkill", "/F", "/IM", proc_name],
                    creationflags=0x08000000,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
                self._log(f"[终止] 已尝试关闭 {proc_name}")
            except Exception:
                pass

    def _terminate_checkin(self):
        """
        处理“终止打卡”按钮点击事件。
        """
        if not self.is_checkin_running:
            self._log("[终止] 当前没有正在运行的打卡任务")
            # 确保按钮状态正确
            self.btn_immediate.config(text="立即打卡", command=self._immediate_checkin,
                                     bg="#4CAF50", fg="white", state="normal")
            return
        
        self._log("[用户操作] >>> 请求强制终止打卡 <<<")
        
        # 1. 立即更新 UI 为“处理中”，防止重复点击
        self.btn_immediate.config(text="终止中...", state="disabled", bg="#999999", fg="white")
        self.root.update_idletasks() # 强制刷新UI
        
        # 2. 设置内部标志位
        self.stop_timer_flag = True 
        self.is_checkin_running = False 
        
        # 3. ✅ 关键：通知打卡模块内部中断
        try:
            import 打卡并发消息 as punch_module
            if hasattr(punch_module, 'CHECKIN_INTERRUPT_EVENT'):
                punch_module.CHECKIN_INTERRUPT_EVENT.set()
                self._log("[终止] 已发送中断信号给打卡模块")
        except Exception as e:
            self._log(f"[终止] 发送中断信号失败: {e}")

        # 4. 执行进程查杀 (辅助手段，防止模块内部阻塞无法退出)
       
        
        # 5. 记录状态，防止今日自动补打
        today_str = datetime.now().strftime("%Y-%m-%d")
        self.last_aborted_date = today_str
        self.today_checkin_attempted = today_str
        self._log(f"[策略] 已记录今日({today_str})手动放弃，今日内不再自动触发补打")
        
        # 6. ✅ 核心修复：直接暴力重置按钮，不依赖任何判断逻辑
        # 延迟稍微长一点（1秒），确保后台线程有机会捕获到中断并开始退出流程
        def _force_reset_button():
            try:
                self.btn_immediate.config(text="立即打卡", command=self._immediate_checkin,
                                         bg="#4CAF50", fg="white", state="normal")
                self._update_status("已终止")
                self._log("[UI] 按钮已强制重置")
            except Exception as e:
                self._log(f"[UI] 重置按钮异常: {e}")

        self.root.after(1000, _force_reset_button)
        
        # 7. 重置停止标志 (延迟重置，防止定时循环立即再次触发)
        def _reset_flag():
            self.stop_timer_flag = False
            
        self.root.after(2000, _reset_flag)
        
        # 6. 恢复 UI 状态
        def _restore_ui():
            self._update_status("已终止")
            self._toggle_immediate_button(to_terminate_mode=False)
            if self.is_timer_running:
                self._update_power_status()
                
        self.root.after(500, _restore_ui) # 稍微延迟一点，确保清理完成
        
        # 7. 重置停止标志
        self.stop_timer_flag = False 

    def _wait_for_desktop_ready(self, timeout=15):
        """
        等待桌面完全就绪（极致稳定版）。
        策略：
        1. 复用 unlock_module.is_lock_screen_active() 确保锁屏判断标准绝对一致。
        2. 确认 Explorer.exe 进程存在且稳定运行。
        3. 给予额外的缓冲时间让桌面图标和任务栏渲染完毕。
        """
        import 亮屏进入桌面 as unlock_module
        
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            # 1. 响应用户的中断请求
            if self.stop_timer_flag:
                return False
                
            # 2. 检查是否还处在锁屏/登录界面 (复用统一接口)
            try:
                if unlock_module.is_lock_screen_active():
                    # 如果还在锁屏，短暂等待后继续循环检测
                    time.sleep(0.5)
                    continue
            except Exception as e:
                # 防止检测函数本身报错导致流程中断，记录日志并继续尝试
                self._log(f"[桌面检测] 锁屏状态检测异常: {e}")
                time.sleep(0.5)
                continue

            # 3. 检查 Explorer 进程是否存在 (桌面资源管理器就绪的核心标志)
            try:
                # 使用 tasklist 快速检查 explorer.exe
                # creationflags=0x08000000 隐藏黑框
                result = subprocess.run(
                    ["tasklist", "/FI", "IMAGENAME eq explorer.exe", "/NH", "/FO", "CSV"],
                    capture_output=True, 
                    text=True, 
                    creationflags=0x08000000
                )
                
                # 检查输出中是否包含 explorer.exe
                if "explorer.exe" in result.stdout.lower():
                    # ✅ 关键优化：Explorer 存在后，再额外等待 0.5~1.0 秒
                    # 这能确保桌面图标、任务栏、系统托盘完全加载完毕，
                    # 防止小程序启动时因为桌面UI未渲染完成而获取焦点失败。
                    time.sleep(0.8) 
                    return True
            except Exception as e:
                self._log(f"[桌面检测] Explorer进程检测异常: {e}")
                pass
            
            # 每次循环间隔，避免CPU占用过高
            time.sleep(0.5)
            
        # 超时仍未就绪
        return False

    def _run_checkin(self, is_timer_task=False, restore_power_after=True):
        try:
            if self.stop_timer_flag or self.is_closing:
                return
            
            # ✅ 优化1：如果是定时任务，先调用亮屏
            if is_timer_task:
                self._log("[系统] 调用亮屏模块...")
                try:
                    import 亮屏进入桌面 as unlock_module
                    # ... (亮屏代码保持不变) ...
                    unlock_module.send_mouse_move(1, 0)
                    time.sleep(0.05)
                    unlock_module.send_mouse_move(-1, 0)
                    time.sleep(0.05)
                    unlock_module.send_mouse_move(1, 0)
                    
                    if unlock_module.is_lock_screen_active():
                        for i in range(3):
                            if self.stop_timer_flag: return 
                            unlock_module.send_key(unlock_module.VK_RETURN, press=True)
                            time.sleep(0.2)
                            unlock_module.send_key(unlock_module.VK_RETURN, press=False)
                            time.sleep(0.8)
                            if not unlock_module.is_lock_screen_active():
                                unlock_module.activate_desktop()
                                break
                    else:
                        unlock_module.activate_desktop()
                    # ==================== 【核心修复】新增：等待桌面完全就绪 ====================
                    self._log("[系统] 亮屏指令已发送，正在等待桌面完全就绪...")
                    if not self._wait_for_desktop_ready(timeout=15):
                        self._log("[警告] 等待桌面就绪超时，可能导致后续操作失败")
                    else:
                        self._log("[系统] 桌面已就绪，开始执行打卡")
                    # ========================================================================

                    self._log("[系统] 亮屏解锁完成")
                except Exception as e:
                    self._log(f"[系统] 亮屏模块异常: {e}")
            


            enable_notify = "1" if self.enable_wechat_notify.get() else "0"
            close_wechat = "0"
            
            self._log("[系统] 调用打卡模块...")
            
            # ✅ 关键：调用打卡模块
            import 打卡并发消息 as punch_module
            
            # 1. 设置回调函数，将打卡模块的 log() 输出转发到主界面
            punch_module.set_log_callback(self._log_to_gui_safe)
            
            # 2. 执行打卡任务
            success, status, message = punch_module.run_full_checkin_task(enable_notify, close_wechat)
            
            # ✅ 核心修复：如果在中途被终止，status 会是 "中断"
            # 此时我们不应该显示弹窗，也不应该执行复杂的 UI 恢复，直接退出即可
            if status == "中断":
                self._log("[系统] 打卡任务已被用户终止，线程安全退出")
                return 

            # ✅ 只有成功或非中断失败时，才更新 UI
            def _restore_focus_and_show_result():
                try:
                    if self.is_closing: return # 如果程序正在关闭，不执行任何UI操作
                    
                    # ✅ 新增：执行微信后续操作
                    action = self.wechat_action_var.get()
                    
                    if action == "关闭窗口":
                        self._log("[微信操作] 正在扫描并最小化所有微信窗口...")
                        minimized_count = 0
                        wechat_pids = set()

                        # --- 步骤 1: 尝试通过多种进程名获取 PID ---
                        target_names = ["WeChat.exe", "Weixin.exe", "WXWork.exe"]
                        
                        for name in target_names:
                            try:
                                # 使用 tasklist 获取 PID
                                output = subprocess.check_output(
                                    ["tasklist", "/FI", f"IMAGENAME eq {name}", "/FO", "CSV", "/NH"],
                                    creationflags=0x08000000
                                ).decode('gbk', errors='ignore').strip()
                                
                                if output:
                                    for line in output.splitlines():
                                        parts = line.split(',')
                                        if len(parts) >= 2:
                                            try:
                                                pid = int(parts[1].strip('"'))
                                                wechat_pids.add(pid)
                                            except:
                                                pass
                            except Exception:
                                pass

                        # 如果 tasklist 没找到，尝试 wmic (更强力)
                        if not wechat_pids:
                            try:
                                output = subprocess.check_output(
                                    ["wmic", "process", "where", "name='WeChat.exe' or name='Weixin.exe'", "get", "ProcessId"],
                                    creationflags=0x08000000
                                ).decode('gbk', errors='ignore')
                                for line in output.splitlines():
                                    line = line.strip()
                                    if line.isdigit():
                                        wechat_pids.add(int(line))
                            except Exception:
                                pass

                        if wechat_pids:
                            self._log(f"[微信操作] 发现 {len(wechat_pids)} 个微信相关进程 ID: {wechat_pids}")
                        else:
                            self._log("[微信操作] ⚠️ 未通过进程名找到微信，尝试通过窗口标题模糊匹配...")

                        # --- 步骤 2: 枚举窗口并最小化 ---
                        def enum_windows_callback(hwnd, extra):
                            nonlocal minimized_count
                            if not win32gui.IsWindowVisible(hwnd):
                                return

                            # 策略 A: 如果找到了 PID，检查窗口是否属于这些 PID
                            if wechat_pids:
                                try:
                                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                                    if pid not in wechat_pids:
                                        return # 不属于微信进程，跳过
                                except:
                                    return # 获取 PID 失败，跳过

                            # 策略 B: 如果没有找到 PID，或者为了双重保险，检查标题和类名
                            title = win32gui.GetWindowText(hwnd)
                            class_name = win32gui.GetClassName(hwnd)
                            
                            # 常见的微信窗口特征
                            is_wechat_window = False
                            
                            # 1. 类名匹配 (覆盖大多数版本)
                            if any(kw in class_name for kw in ["WeChat", "TXGui", "Chrome_WidgetWin_1"]): # Chrome_WidgetWin_1 是新版微信内核
                                is_wechat_window = True
                            
                            # 2. 标题匹配 (如果类名不确定，但标题包含微信)
                            if not is_wechat_window and title:
                                if any(kw in title for kw in ["微信", "WeChat", "文件传输助手"]):
                                    is_wechat_window = True

                            if is_wechat_window:
                                try:
                                    # 排除掉一些明显的非主窗口（如小的弹窗、菜单等，通常没有标题或很小）
                                    # 这里我们主要想最小化主界面，所以如果有标题且比较大，或者是主类名，就最小化
                                    if title or "WeChatMainWnd" in class_name or "TXGuiFoundation" in class_name:
                                        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                                        win32api.PostMessage(hwnd, win32con.WM_SYSCOMMAND, 0xF020, 0)
                                        minimized_count += 1
                                        self._log(f"[微信操作] 已最小化: '{title}'")
                                except Exception as e:
                                    pass

                        win32gui.EnumWindows(enum_windows_callback, None)

                        if minimized_count > 0:
                            self._log(f"[微信操作] ✅ 成功最小化了 {minimized_count} 个窗口")
                        else:
                            self._log("[微信操作] ⚠️ 未找到可最小化的微信窗口 (可能已最小化或未运行)")

                    elif action == "退出微信":
                        self._log("[微信操作] 尝试强制退出微信进程...")
                        try:
                            # 尝试杀死所有可能的微信进程名
                            for name in ["WeChat.exe", "Weixin.exe", "WXWork.exe"]:
                                subprocess.run(
                                    ["taskkill", "/F", "/IM", name],
                                    creationflags=0x08000000,
                                    stdout=subprocess.DEVNULL,
                                    stderr=subprocess.DEVNULL
                                )
                            self._log("[微信操作] 已发送退出指令")
                        except Exception as e:
                            self._log(f"[微信操作] 退出失败: {e}")
                    

                    # ✅ 修复闪烁：优化置顶和弹窗逻辑
                    # 1. 确保主窗口在最前面并获得焦点，作为弹窗的锚点
                    self.root.lift()
                    self.root.focus_force()
                    
                    # 2. 关键：强制处理所有待处理的绘图事件，防止“半成品”窗口闪现
                    self.root.update_idletasks()
                    
                    # 3. 设置主窗口临时置顶，确保它比桌面和其他背景窗口高
                    self.root.attributes('-topmost', True)
                    
                    # 4. 极短等待，让 Windows 系统完成窗口层级的重排
                    time.sleep(0.1) 
                    
                    # 5. 显示结果弹窗
                    if success:
                        self._update_status("打卡成功")
                        if not is_timer_task:
                            # AutoCloseMessageBox 内部已经处理了置顶和居中
                            AutoCloseMessageBox(self.root, "打卡成功", "打卡完成", auto_close_seconds=5)
                    else:
                        self._update_status(f"打卡{status}")
                        if not is_timer_task:
                            AutoCloseMessageBox(self.root, f"打卡{status}", message, auto_close_seconds=5, is_error=True)
                    
                    # 6. 弹窗关闭后（或显示后），取消主窗口的强制置顶，恢复正常行为
                    # 注意：AutoCloseMessageBox 是模态的 (grab_set)，代码会卡在这里直到弹窗关闭
                    self.root.after(100, lambda: self.root.attributes('-topmost', False))
                            
                except Exception as e:
                    pass # 忽略UI错误
                finally:
                    # 这里的 finally 原本用于取消置顶，现在移到了弹窗逻辑之后更合适
                    # 但为了保险，保留一个延迟取消置顶的逻辑
                    if not self.is_closing:
                        self.root.after(200, lambda: self.root.attributes('-topmost', False))

            self.root.after(0, _restore_focus_and_show_result)
                    
        except Exception as e:
            # 如果是中断导致的异常，记录一下即可
            if self.stop_timer_flag:
                self._log("[系统] 捕获到终止相关的异常，忽略")
            else:
                self._log(f"[打卡] 整体异常: {e}")
                self._update_status("打卡异常")
        finally:
            # ✅ 关键：无论成功失败，都要重置运行标志
            today_str = datetime.now().strftime("%Y-%m-%d")
            self.today_checkin_attempted = today_str
            self.is_checkin_running = False
            
            # ✅ 只有在非终止状态下，才尝试恢复按钮
            # 如果是终止状态，_terminate_checkin 已经接管了按钮控制
            if not self.stop_timer_flag:
                 self.root.after(100, lambda: self._toggle_immediate_button(to_terminate_mode=False))
            else:
                 self._log("[线程] 检测到终止标志，跳过按钮重置，交由主线程处理")

    def _start_timer(self, auto_confirm=False):
        if self.is_timer_running:
            return
        if self.is_closing:
            return
        
        if not self.shortcut_detected:
            if not auto_confirm:
                if not self._show_shortcut_warning(allow_continue=True):
                    return
            else:
                self.shortcut_status_label.config(text="目标快捷方式", fg="green")
        
        hour = int(self.schedule_hour.get())
        minute = int(self.schedule_minute.get())
        now = datetime.now()
        target = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if target <= now:
            target += timedelta(days=1)
        
        if not auto_confirm:
            if not messagebox.askyesno("确认", f"是否启动定时打卡？\n下次打卡：{target.strftime('%H:%M')}"):
                return
        
        self._log("\n【电源】正在通过 API 阻止系统睡眠...")
        self.api_sleep_preventer.prevent()

        self.config_manager.set("is_timer_enabled", True)
        self.config_manager.save_config()
        self.last_timer_state = True
        
        self.is_timer_running = True
        self.stop_timer_flag = False
        self.btn_timer_toggle.config(text="停止定时", bg="#ff4848", fg="white")
        self._update_status("定时运行中", target.strftime("%Y-%m-%d %H:%M:%S"))
        self._update_power_status()
        
        self.timer_thread = threading.Thread(target=self._timer_loop, args=(hour, minute), daemon=True)
        self.timer_thread.start()
    
    def _timer_loop(self, hour, minute):
        last_checked_date = None
        # ✅ 新增：记录上次执行打卡的时间戳，用于防止短时间内的重复触发
        last_execution_time = 0 
        
        while not self.is_closing: 
            current_date = datetime.now().strftime("%Y-%m-%d")
            if last_checked_date != current_date:
                if self.last_aborted_date and self.last_aborted_date != current_date:
                    self.last_aborted_date = None
                if self.today_checkin_attempted and self.today_checkin_attempted != current_date:
                    self.today_checkin_attempted = None
                    self._log("[系统] 新的一天，重置打卡状态")
                last_checked_date = current_date

            if self.stop_timer_flag:
                time.sleep(1)
                continue

            # ✅ 1. 动态获取最新时间
            try:
                current_hour = int(self.schedule_hour.get())
                current_minute = int(self.schedule_minute.get())
            except:
                current_hour = hour
                current_minute = minute
            
            # ✅ 2. 处理时间变更标志
            if self.time_config_changed:
                self._log(f"[定时] 检测到时间变更为 {current_hour:02d}:{current_minute:02d}，重新计算目标时间")
                self.time_config_changed = False
                last_checked_date = None # 强制刷新日期逻辑

            # ✅ 3. 计算目标
            should_checkin_now, target = self._calculate_next_target_and_check_missed(current_hour, current_minute)
            
            # 更新 UI
            self.root.after(0, lambda t=target: self.next_time_label.config(text=f"下次打卡：{t.strftime('%Y-%m-%d %H:%M:%S')}"))

            # ✅ 4. 核心逻辑：是否应该立即执行（补打或准时）
            if should_checkin_now:
                enable_grace = self.config_manager.get("enable_grace_checkin", False)
                
                # --- 安全检查 A: 是否正在运行？ ---
                if self.is_checkin_running:
                    # 如果已经在打卡，忽略本次时间变更触发的请求，等待当前任务结束
                    time.sleep(5) 
                    continue
                
                # --- 安全检查 B: 今天是否已经打过/试过？ ---
                if self.today_checkin_attempted == current_date:
                    # 今天已尝试过，不再重复，即使改了时间
                    time.sleep(60)
                    continue

                # --- 安全检查 C: 防抖 (可选) ---
                import time as time_module
                now_ts = time_module.time()
                if now_ts - last_execution_time < 60: # 60秒内不重复执行
                     time.sleep(5)
                     continue

                if not enable_grace:
                    # 如果没开宽限，且时间已过，说明今天没戏了，等明天
                    time.sleep(60)
                    continue
                
                if self.last_aborted_date == current_date:
                    # 用户手动终止过，尊重用户意愿，今天不再自动打
                    time.sleep(60)
                    continue

                # ✅ 通过所有检查，执行打卡
                self._log("[检测] 时间变更或到达，触发立即打卡逻辑")
                self._update_status("正在补打卡...", target.strftime("%Y-%m-%d %H:%M:%S"))
                
                self.is_checkin_running = True
                last_execution_time = time_module.time() # 记录执行时间
                self.root.after(0, lambda: self._toggle_immediate_button(to_terminate_mode=True))
                
                # 注意：_run_checkin 是阻塞在当前线程的吗？不，它是同步执行的，但内部可能有 sleep
                # 为了防止 _run_checkin 执行期间循环卡死，我们最好还是让它在后台跑，但 _timer_loop 本身就在后台线程
                # 所以直接调用是可以的，但要注意 _run_checkin 内部不要有无限循环
                self._run_checkin(is_timer_task=True)
                
                continue

            # 5. 正常等待逻辑
            now = datetime.now()
            delay = (target - now).total_seconds()
            if delay < 0: delay = 1
            
            elapsed = 0
            step = 1 
            while elapsed < delay:
                if self.stop_timer_flag or self.is_closing: break 
                if self.time_config_changed: break # 时间变了，提前醒来
                time.sleep(step)
                elapsed += step
            
            if self.is_closing: break
            if self.stop_timer_flag: continue
            
            # 6. 准时触发前的最后检查
            if self.is_checkin_running:
                self._log("[定时] 到达预定时间，但当前已有任务在运行，跳过")
                continue

            self._log("[定时] 到达预定时间，开始执行打卡")
            self._update_status("正在打卡...")
            
            self.is_checkin_running = True
            last_execution_time = time.time() # 记录执行时间
            self.root.after(0, lambda: self._toggle_immediate_button(to_terminate_mode=True))
            
            self._run_checkin(is_timer_task=True)
    
    def _stop_timer(self):
        if not self.is_timer_running:
            return
        
        self.stop_timer_flag = True
       
        
        self._log("\n【电源】正在恢复系统自动睡眠...")
        self.api_sleep_preventer.allow()
        
        self.is_timer_running = False
        self.btn_timer_toggle.config(text="定时打卡", bg="#2196F3", fg="white")
        self._update_status("已停止")
        self._update_power_status()

        self.config_manager.set("is_timer_enabled", False)
        self.config_manager.save_config()
        self.last_timer_state = False
        
        messagebox.showinfo("提示", "定时任务已停止")
    
    def _on_closing(self):
        # ✅ 修复：优化关闭窗口逻辑，使其更符合用户直觉
        # 逻辑调整：
        # 【是】-> 彻底退出程序
        # 【否】-> 最小化到托盘（后台运行）
        # 【取消】-> 取消操作，返回窗口
        
        # 1. 如果还没有触发过“真正关闭”流程，先询问用户
        if not self.is_closing:
            result = messagebox.askyesnocancel(
                "关闭窗口", 
                "【是】彻底退出程序（停止所有后台任务）\n"
                "【否】最小化到系统托盘（后台继续定时运行）\n"
                "【取消】取消操作（返回主界面）",
                icon='warning' # 使用 warning 图标引起注意
            )
            
            if result is None:
                # 用户点击了“取消”按钮 或 弹窗右上角的“X”
                return
            
            if result:
                # 用户选择“是”：彻底退出
                self.is_closing = True
            else:
                # 用户选择“否”：最小化到托盘
                self._hide_to_tray()
                return

        # 2. 如果已经在关闭流程中，防止重入
        if hasattr(self, '_is_cleaning_up') and self._is_cleaning_up:
            return
        self._is_cleaning_up = True
        
        self._log("【关闭程序】正在清理资源...")

        # 3. 保存配置
        self.config_manager.set("is_timer_enabled", self.is_timer_running)
        self._save_all_config()
        
        # 4. 停止托盘图标
        if hasattr(self, 'tray_icon') and self.tray_icon:
            try:
                self.tray_icon.stop()
            except:
                pass

        # 5. 停止所有后台任务
        self.stop_timer_flag = True
        
        if self.is_timer_running:
            self.is_timer_running = False
            self.api_sleep_preventer.allow() # 确保释放防睡眠
        
        # 6. 短暂等待线程结束
        time.sleep(0.2)
        
        # 7. 释放 Mutex
        if self.h_mutex:
            try:
                win32api.CloseHandle(self.h_mutex)
            except:
                pass

        # 8. 销毁窗口
        try:
            self.root.destroy()
        except:
            pass
    
    def run(self):
        self._create_tray_icon()
        try:
            self.root.mainloop()
        except Exception as e:
            pass
        finally:
            if not self.is_closing:
                self.is_closing = True
                self._on_closing()


def main():
    """主程序入口"""
    app = AutoCheckInGUI()
    app.run()


if __name__ == "__main__":
    main()