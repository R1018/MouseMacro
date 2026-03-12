# 依赖说明：
# 需要安装以下包（建议用管理员权限运行命令行）：
# pip install pywin32 pynput

import json
import time
import threading
import win32gui
import win32con
from pynput import mouse
from pynput.mouse import Controller, Button
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import ctypes
import os
import sys
from pathlib import Path

# 设置 DPI 感知
ctypes.windll.user32.SetProcessDPIAware()

# 隐藏控制台窗口
if sys.platform == 'win32':
    try:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)
    except:
        pass

# 配置文件路径
CONFIG_FILE = "config.json"

# 全局状态
recorder_instance = None
player_instance = None

# 日志控制变量
log_to_console = [True]  # 用列表包裹以便在闭包中修改

def debug_log(msg):
    # 精简日志格式
    log_msg = f"{time.strftime('%Y-%m-%d %H:%M:%S')} {msg}\n"

    # 检查日志文件大小，超过 5MB 则清空
    log_file = Path("debug.log")
    if log_file.exists() and log_file.stat().st_size > 5 * 1024 * 1024:
        log_file.write_text(log_msg, encoding="utf-8")
    else:
        with open("debug.log", "a", encoding="utf-8") as f:
            f.write(log_msg)

    if log_to_console[0]:
        print(log_msg, end="")

# 配置管理
def load_config():
    """加载配置文件"""
    default_config = {
        'theme': 'auto',
        'last_script': '',
        'last_window': '',
        'log_enabled': True
    }
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                default_config.update(config)
    except:
        pass
    return default_config

def save_config(config):
    """保存配置文件"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except:
        pass

# 检查窗口是否存在
def check_window_exists(window_title):
    """检查指定窗口是否存在"""
    try:
        hwnd = win32gui.FindWindow(None, window_title)
        return hwnd != 0
    except:
        return False

class MouseRecorder:
    def __init__(self, window_title, drag_interval=0.01):
        self.window_title = window_title
        self.drag_interval = drag_interval  # 拖动时的记录间隔
        self.events = []
        self.start_time = None
        self.initial_rect = None
        self.is_dragging = False
        self.last_drag_time = 0
        self.listener = None
    
    def _get_window_rect(self):
        hwnd = win32gui.FindWindow(None, self.window_title)
        if hwnd == 0:
            raise ValueError(f"窗口 '{self.window_title}' 未找到")
        left_top = win32gui.ClientToScreen(hwnd, (0, 0))
        right_bottom = win32gui.ClientToScreen(hwnd, win32gui.GetClientRect(hwnd)[2:])
        return (left_top[0], left_top[1], right_bottom[0], right_bottom[1])
    
    def _to_relative(self, x, y):
        left, top, right, bottom = self.initial_rect
        width = max(1, right - left)
        height = max(1, bottom - top)
        rel_x = max(0.0, min(1.0, (x - left) / width))
        rel_y = max(0.0, min(1.0, (y - top) / height))
        return rel_x, rel_y
    
    def _is_in_target_window(self, x, y):
        try:
            hwnd = win32gui.FindWindow(None, self.window_title)
            if hwnd == 0:
                return False
            left_top = win32gui.ClientToScreen(hwnd, (0, 0))
            right_bottom = win32gui.ClientToScreen(hwnd, win32gui.GetClientRect(hwnd)[2:])
            if left_top[0] <= x <= right_bottom[0] and left_top[1] <= y <= right_bottom[1]:
                return True
            return False
        except Exception as e:
            debug_log(f"_is_in_target_window error: {e}")
            return False
    
    def _on_move(self, x, y):
        if not self.start_time:
            return
        if not self._is_in_target_window(x, y):
            return
        current_time = time.time()
        rel_x, rel_y = self._to_relative(x, y)
        event_time = current_time - self.start_time
        if self.is_dragging and (current_time - self.last_drag_time >= self.drag_interval):
            debug_log(f"Move: {rel_x:.3f},{rel_y:.3f} t={event_time:.2f}")
            self.events.append({
                'type': 'move',
                'x': rel_x,
                'y': rel_y,
                'time': event_time
            })
            self.last_drag_time = current_time

    def _on_click(self, x, y, button, pressed):
        if not self.start_time:
            return
        event_time = time.time() - self.start_time
        if pressed:
            if not self._is_in_target_window(x, y):
                return
            rel_x, rel_y = self._to_relative(x, y)
            debug_log(f"Click: {rel_x:.3f},{rel_y:.3f} {button.name} down t={event_time:.2f}")
            self.is_dragging = True
            self.last_drag_time = time.time()
            self.events.append({
                'type': 'click',
                'x': rel_x,
                'y': rel_y,
                'button': button.name,
                'state': 'pressed',
                'time': event_time
            })
        else:
            try:
                rel_x, rel_y = self._to_relative(x, y)
            except Exception:
                rel_x, rel_y = x, y
            debug_log(f"Click: {rel_x:.3f},{rel_y:.3f} {button.name} up t={event_time:.2f}")
            self.is_dragging = False
            self.events.append({
                'type': 'click',
                'x': rel_x,
                'y': rel_y,
                'button': button.name,
                'state': 'released',
                'time': event_time
            })

    def start_recording(self):
        self.initial_rect = self._get_window_rect()
        self.events = []
        self.start_time = time.time()
        self.is_dragging = False
        # 覆盖debug.log
        with open("debug.log", "w", encoding="utf-8") as f:
            f.write("")
        debug_log(f"Start recording: {self.window_title} rect={self.initial_rect}")
        self.listener = mouse.Listener(
            on_move=self._on_move,
            on_click=self._on_click
        )
        self.listener.start()
        debug_log(f"Recording started...")

    def stop_recording(self):
        if self.listener:
            self.listener.stop()
        self.start_time = None
        debug_log(f"Stop recording. Events: {len(self.events)}")

    def save_recording(self, filename):
        debug_log(f"Save recording. Events: {len(self.events)}")
        if not self.events:
            debug_log("No data to save")
            return
        data = {
            'window_title': self.window_title,
            'initial_rect': self.initial_rect,
            'drag_interval': self.drag_interval,
            'events': self.events
        }
        with open(filename, 'w') as f:
            json.dump(data, f)
        debug_log(f"Saved to {filename}")

class MousePlayer:
    def __init__(self, filename):
        with open(filename, 'r') as f:
            self.data = json.load(f)
        self.mouse = Controller()
        self.is_playing = False
        self.play_thread = None
    
    def _get_current_rect(self):
        """获取当前窗口内容区的位置和大小（不含边框和标题栏）"""
        hwnd = win32gui.FindWindow(None, self.data['window_title'])
        if hwnd == 0:
            raise ValueError(f"窗口 '{self.data['window_title']}' 未找到")
        
        # 如果窗口最小化则恢复
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.2)
        
        # 确保窗口在前台
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.1)
        
        # 获取内容区左上和右下坐标
        left_top = win32gui.ClientToScreen(hwnd, (0, 0))
        right_bottom = win32gui.ClientToScreen(hwnd, win32gui.GetClientRect(hwnd)[2:])
        return (left_top[0], left_top[1], right_bottom[0], right_bottom[1])
    
    def _to_absolute(self, rel_x, rel_y, rect):
        """将相对坐标转换为绝对坐标"""
        left, top, right, bottom = rect
        width = max(1, right - left)
        height = max(1, bottom - top)
        x = left + int(rel_x * width)
        y = top + int(rel_y * height)
        return x, y

    def _play_events(self, times=1):
        try:
            for i in range(times):
                # 每次循环开始前获取一次窗口位置
                current_rect = self._get_current_rect()
                prev_time = 0
                for event in self.data['events']:
                    if not self.is_playing:
                        break
                    # 等待到事件发生的时间点
                    wait_time = event['time'] - prev_time
                    if wait_time > 0:
                        time.sleep(wait_time)
                    prev_time = event['time']

                    # 转换坐标
                    x, y = self._to_absolute(event['x'], event['y'], current_rect)

                    if event['type'] == 'move':
                        self.mouse.position = (x, y)
                    elif event['type'] == 'click':
                        # 对于点击事件，先移动再点击
                        self.mouse.position = (x, y)
                        button = Button[event['button']]
                        if event['state'] == 'pressed':
                            self.mouse.press(button)
                        else:
                            self.mouse.release(button)
        finally:
            self.is_playing = False

    def play(self, times=1, blocking=False):
        """回放录制的鼠标操作"""
        if self.is_playing:
            print("已经在回放中")
            return
        self.is_playing = True
        self.play_thread = threading.Thread(target=lambda: self._play_events(times))
        self.play_thread.daemon = True
        self.play_thread.start()
        if blocking:
            self.play_thread.join()
    
    def stop(self):
        """停止回放"""
        self.is_playing = False
        if self.play_thread and self.play_thread.is_alive():
            self.play_thread.join(timeout=0.5)

def gui_main():
    # 加载配置
    config = load_config()
    log_to_console[0] = config.get('log_enabled', True)

    # 主题配置
    themes = {
        'light': {
            'bg': '#f0f0f0',
            'frame_bg': '#f0f0f0',
            'entry_bg': '#fff',
            'entry_fg': '#000',
            'label_fg': '#000',
            'btn_bg': '#ddd',
            'btn_fg': '#000',
            'btn_active_bg': '#ccc',
            'border_color': '#999'
        },
        'dark': {
            'bg': '#111',
            'frame_bg': '#111',
            'entry_bg': '#222',
            'entry_fg': '#fff',
            'label_fg': '#fff',
            'btn_bg': '#333',
            'btn_fg': '#fff',
            'btn_active_bg': '#555',
            'border_color': '#fff'
        }
    }

    # 检测系统主题
    def get_system_theme():
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                r'Software\Microsoft\Windows\CurrentVersion\Themes\Personalize')
            value, _ = winreg.QueryValueEx(key, 'AppsUseLightTheme')
            winreg.CloseKey(key)
            return 'light' if value == 1 else 'dark'
        except:
            return 'dark'

    # 当前主题变量
    saved_theme = config.get('theme', 'auto')

    root = tk.Tk()
    root.title("MouseMacro - 鼠标操作录制工具")
    root.geometry("600x620")
    root.minsize(550, 580)

    # 在创建 root 之后才创建 StringVar
    current_theme = tk.StringVar(value='dark')

    # 应用主题
    def apply_theme(theme_name):
        theme = themes[theme_name]
        root.configure(bg=theme['bg'])

        # 更新主题切换和导航区域的背景
        try:
            theme_frame.configure(bg=theme['bg'])
            nav_frame.configure(bg=theme['bg'])
        except:
            pass

        # 更新所有控件
        def update_widget(widget):
            try:
                if isinstance(widget, (tk.Label, tk.Button)):
                    # 跳过特定按钮（它们的背景色单独处理）
                    widget_name = str(widget)
                    if '停止' in widget.cget('text') or '开始' in widget.cget('text'):
                        # 不改变操作按钮的背景色
                        widget.configure(fg=theme['label_fg'])
                    else:
                        widget.configure(bg=theme.get('frame_bg', theme['bg']),
                                        fg=theme['label_fg'])
                elif isinstance(widget, tk.Entry):
                    widget.configure(bg=theme['entry_bg'],
                                    fg=theme['entry_fg'],
                                    insertbackground=theme['entry_fg'])
                elif isinstance(widget, tk.Frame):
                    widget.configure(bg=theme.get('frame_bg', theme['bg']))
            except:
                pass
            for child in widget.winfo_children():
                update_widget(child)

        update_widget(root)

    # 主题切换
    def toggle_theme(theme_name):
        current_theme.set(theme_name)
        apply_theme(theme_name)

    def auto_theme():
        theme = get_system_theme()
        toggle_theme(theme)

    # 初始化主题（稍后应用）
    initial_theme = get_system_theme() if saved_theme == 'auto' else saved_theme
    current_theme.set(initial_theme)

    # 通用样式（动态主题）
    style_font = ("微软雅黑", 12)

    def get_theme_colors():
        theme_name = current_theme.get()
        return themes[theme_name]

    # 页面框架
    record_frame = tk.Frame(root)
    play_frame = tk.Frame(root)
    for frame in (record_frame, play_frame):
        frame.grid(row=2, column=0, sticky="nsew")
        frame.grid_rowconfigure(99, weight=1)
        frame.grid_columnconfigure(0, weight=1)

    # 切换显示的框架
    def show_frame(f):
        f.tkraise()
        # 刷新主题
        apply_theme(current_theme.get())

    # 辅助函数：创建主题感知的控件
    def make_btn(parent, text, cmd, bg=None):
        colors = get_theme_colors()
        if bg is None:
            bg = colors['btn_bg']

        # 根据主题设置不同的边框颜色
        border_color = '#999' if current_theme.get() == 'light' else '#666'

        return tk.Button(
            parent, text=text, command=cmd,
            font=style_font, bg=bg, fg=colors['btn_fg'],
            activebackground=colors['btn_active_bg'],
            activeforeground=colors['btn_fg'],
            relief="raised", bd=2, padx=10, pady=2
        )

    def make_label(parent, text):
        colors = get_theme_colors()
        return tk.Label(parent, text=text, font=style_font,
                       fg=colors['label_fg'], anchor="w")

    def make_entry(parent, var):
        colors = get_theme_colors()
        return tk.Entry(parent, textvariable=var, font=style_font,
                       bg=colors['entry_bg'], fg=colors['entry_fg'],
                       insertbackground=colors['entry_fg'],
                       relief="flat", bd=2)

    # 变量
    scripts_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Scripts")
    os.makedirs(scripts_dir, exist_ok=True)
    default_recording_file = os.path.join(scripts_dir, "mouse_recording.json")
    recording_file = tk.StringVar(value=config.get('last_script', default_recording_file))
    drag_interval = tk.DoubleVar(value=0.01)
    window_title = tk.StringVar(value=config.get('last_window', ""))
    playback_times = tk.IntVar(value=1)

    # 状态标签
    status_var = tk.StringVar(value="就绪")

    # 录制/回放状态
    is_recording = tk.BooleanVar(value=False)
    is_playing = tk.BooleanVar(value=False)

    # 录制页面
    make_label(record_frame, "窗口标题:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
    make_entry(record_frame, window_title).grid(row=1, column=0, sticky="ew", padx=10)
    make_btn(record_frame, "选择窗口", lambda: select_window_title(root, window_title)).grid(row=2, column=0, sticky="ew", padx=10, pady=2)
    make_btn(record_frame, "检查窗口", lambda: verify_window(window_title, status_var)).grid(row=3, column=0, sticky="ew", padx=10, pady=2)
    make_label(record_frame, "拖动记录间隔(秒):").grid(row=4, column=0, sticky="w", padx=10, pady=5)
    make_entry(record_frame, drag_interval).grid(row=5, column=0, sticky="ew", padx=10)
    make_label(record_frame, "录制文件:").grid(row=6, column=0, sticky="w", padx=10, pady=5)
    make_entry(record_frame, recording_file).grid(row=7, column=0, sticky="ew", padx=10)
    make_btn(record_frame, "选择文件", lambda: select_recording_file(recording_file, scripts_dir)).grid(row=8, column=0, sticky="ew", padx=10, pady=2)

    # 录制按钮组
    record_btn_frame = tk.Frame(record_frame)
    record_btn_frame.grid(row=9, column=0, sticky="ew", padx=10, pady=5)

    btn_record = make_btn(record_btn_frame, "开始录制",
                         lambda: threading.Thread(target=lambda: start_recording(window_title, drag_interval, recording_file, root, scripts_dir, is_recording, status_var, config), daemon=True).start(),
                         bg="#444")
    btn_record.pack(side="left", fill="x", expand=True, padx=(0, 2))

    btn_stop_record = make_btn(record_btn_frame, "停止录制", stop_recording, bg="#664")
    btn_stop_record.pack(side="left", fill="x", expand=True, padx=(2, 0))

    # 状态显示
    status_label = tk.Label(record_frame, textvariable=status_var, font=("微软雅黑", 10))
    status_label.grid(row=10, column=0, sticky="ew", padx=10, pady=5)

    # 回放页面
    make_label(play_frame, "录制文件:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
    make_entry(play_frame, recording_file).grid(row=1, column=0, sticky="ew", padx=10)
    make_btn(play_frame, "选择文件", lambda: select_recording_file(recording_file, scripts_dir)).grid(row=2, column=0, sticky="ew", padx=10, pady=2)
    make_btn(play_frame, "刷新脚本列表", lambda: refresh_scripts(scripts_dir, recording_file)).grid(row=3, column=0, sticky="ew", padx=10, pady=2)
    make_label(play_frame, "回放次数:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
    make_entry(play_frame, playback_times).grid(row=5, column=0, sticky="ew", padx=10)

    # 回放按钮组
    play_btn_frame = tk.Frame(play_frame)
    play_btn_frame.grid(row=6, column=0, sticky="ew", padx=10, pady=5)

    btn_play = make_btn(play_btn_frame, "开始回放",
                       lambda: threading.Thread(target=lambda: start_playback(recording_file, playback_times, root, is_playing, status_var, config), daemon=True).start(),
                       bg="#444")
    btn_play.pack(side="left", fill="x", expand=True, padx=(0, 2))

    btn_stop_play = make_btn(play_btn_frame, "停止回放", stop_playback, bg="#664")
    btn_stop_play.pack(side="left", fill="x", expand=True, padx=(2, 0))

    # 状态显示
    play_status_label = tk.Label(play_frame, textvariable=status_var, font=("微软雅黑", 10))
    play_status_label.grid(row=7, column=0, sticky="ew", padx=10, pady=5)

    btn_exit = make_btn(play_frame, "退出", lambda: on_exit(root, config), bg="#222")
    btn_exit.grid(row=99, column=0, sticky="ew", padx=10, pady=10)

    # 主题切换按钮区域
    theme_frame = tk.Frame(root, bg=themes[current_theme.get()]['bg'])
    theme_frame.grid(row=0, column=0, sticky="ew")

    def make_theme_btn(text, cmd):
        colors = get_theme_colors()
        border = '#666' if current_theme.get() == 'light' else '#888'
        return tk.Button(
            theme_frame, text=text, command=cmd,
            font=("Segoe UI Emoji", 12),
            bg=colors['btn_bg'], fg=colors['btn_fg'],
            activebackground=colors['btn_active_bg'],
            relief="raised", bd=3, width=4, padx=5, pady=3
        )

    make_theme_btn("☀", lambda: toggle_theme('light')).pack(side="left", padx=3, pady=8)
    make_theme_btn("☾", lambda: toggle_theme('dark')).pack(side="left", padx=3, pady=8)
    make_theme_btn("🔄", auto_theme).pack(side="left", padx=3, pady=8)

    # 日志输出控制按钮
    def toggle_log():
        log_to_console[0] = not log_to_console[0]
        btn_log.config(text=f"日志输出: {'开' if log_to_console[0] else '关'}")

    colors = get_theme_colors()
    btn_log = tk.Button(
        theme_frame,
        text=f"日志输出: {'开' if log_to_console[0] else '关'}",
        command=toggle_log,
        font=style_font,
        bg=colors['btn_bg'], fg=colors['btn_fg'],
        activebackground=colors['btn_active_bg'],
        relief="raised", bd=2, padx=8, pady=2
    )
    btn_log.pack(side="right", padx=8, pady=8)

    # 页面切换按钮
    nav_frame = tk.Frame(root, bg=themes[current_theme.get()]['bg'])
    nav_frame.grid(row=1, column=0, sticky="ew")

    nav_colors = get_theme_colors()
    tk.Button(
        nav_frame, text="录制",
        command=lambda: show_frame(record_frame),
        font=("微软雅黑", 11, "bold"),
        bg=nav_colors['btn_bg'], fg=nav_colors['btn_fg'],
        activebackground=nav_colors['btn_active_bg'],
        relief="raised", bd=3
    ).pack(side="left", fill="x", expand=True, padx=3, pady=6)

    tk.Button(
        nav_frame, text="回放",
        command=lambda: show_frame(play_frame),
        font=("微软雅黑", 11, "bold"),
        bg=nav_colors['btn_bg'], fg=nav_colors['btn_fg'],
        activebackground=nav_colors['btn_active_bg'],
        relief="raised", bd=3
    ).pack(side="left", fill="x", expand=True, padx=3, pady=6)

    root.grid_rowconfigure(2, weight=1)
    root.grid_columnconfigure(0, weight=1)

    print("正在显示初始界面...")
    show_frame(record_frame)
    root.mainloop()

# 验证窗口是否存在
def verify_window(window_title_var, status_var):
    title = window_title_var.get()
    if not title:
        status_var.set("错误: 请先输入窗口标题")
        return
    if check_window_exists(title):
        status_var.set(f"验证通过: 窗口 '{title}' 存在")
    else:
        status_var.set(f"错误: 窗口 '{title}' 不存在")

# 刷新脚本列表
def refresh_scripts(scripts_dir, recording_file_var):
    """显示脚本选择对话框"""
    scripts = []
    try:
        for f in os.listdir(scripts_dir):
            if f.endswith('.json'):
                scripts.append(os.path.join(scripts_dir, f))
    except:
        pass

    if not scripts:
        messagebox.showinfo("提示", "Scripts 目录中没有找到脚本文件")
        return

    # 创建选择对话框
    dialog = tk.Toplevel()
    dialog.title("选择脚本")
    dialog.geometry("600x400")
    dialog.transient(dialog.master)
    dialog.grab_set()

    # 脚本列表
    list_frame = tk.Frame(dialog)
    list_frame.pack(fill="both", expand=True, padx=10, pady=10)

    scrollbar = tk.Scrollbar(list_frame)
    scrollbar.pack(side="right", fill="y")

    listbox = tk.Listbox(list_frame, font=("微软雅黑", 10), yscrollcommand=scrollbar.set)
    listbox.pack(fill="both", expand=True)
    scrollbar.config(command=listbox.yview)

    # 填充列表
    for script in scripts:
        listbox.insert(tk.END, os.path.basename(script))

    # 双击选择
    def on_select(event=None):
        selection = listbox.curselection()
        if selection:
            selected = scripts[selection[0]]
            recording_file_var.set(selected)
            dialog.destroy()

    listbox.bind("<Double-Button-1>", on_select)

    # 确认按钮
    btn_frame = tk.Frame(dialog)
    btn_frame.pack(fill="x", padx=10, pady=(0, 10))

    tk.Button(btn_frame, text="确定", command=on_select).pack(side="right", padx=5)
    tk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side="right", padx=5)

# 停止录制
def stop_recording():
    global recorder_instance
    if recorder_instance:
        try:
            recorder_instance.stop_recording()
        except:
            pass

# 停止回放
def stop_playback():
    global player_instance
    if player_instance:
        try:
            player_instance.stop()
        except:
            pass

# 退出时保存配置
def on_exit(root, config):
    save_config(config)
    root.quit()

# 保存配置辅助函数
def update_config(config, key, value):
    config[key] = value
    save_config(config)

# 选择录制文件时只允许 Scripts 目录下
def select_recording_file(recording_file, scripts_dir):
    from tkinter import filedialog
    filename = filedialog.asksaveasfilename(
        title="选择录制文件",
        initialdir=scripts_dir,
        filetypes=[("JSON文件", "*.json")],
        defaultextension=".json"
    )
    if filename:
        recording_file.set(filename)

def start_recording(window_title, drag_interval, recording_file, root, scripts_dir, is_recording, status_var, config):
    global recorder_instance
    from tkinter import messagebox, simpledialog

    if is_recording.get():
        messagebox.showwarning("提示", "正在录制中，请先停止录制")
        return

    wt = window_title.get()
    if not wt:
        messagebox.showerror("错误", "窗口标题不能为空", parent=root)
        return

    # 验证窗口是否存在
    if not check_window_exists(wt):
        messagebox.showerror("错误", f"窗口 '{wt}' 不存在或不可见", parent=root)
        status_var.set(f"错误: 窗口 '{wt}' 不存在")
        return

    window_title.set(wt)
    interval = drag_interval.get()
    filename = recording_file.get() or os.path.join(scripts_dir, "mouse_recording.json")

    # 检查重名
    if os.path.exists(filename):
        res = messagebox.askyesno("文件已存在", f"文件 {os.path.basename(filename)} 已存在，是否覆盖？", parent=root)
        if not res:
            return

    try:
        is_recording.set(True)
        status_var.set("录制中... 按 ESC 键停止")

        # 更新配置
        config['last_window'] = wt
        config['last_script'] = filename
        save_config(config)

        recorder = MouseRecorder(wt, interval)
        recorder_instance = recorder
        recorder.start_recording()

        # 使用非阻塞方式监听键盘
        from pynput import keyboard
        stop_event = threading.Event()

        def on_press(key):
            if key == keyboard.Key.esc:
                stop_event.set()
                return False

        # 创建键盘监听器
        keyboard_listener = keyboard.Listener(on_press=on_press)
        keyboard_listener.start()

        # 等待停止信号
        while not stop_event.is_set() and is_recording.get():
            time.sleep(0.1)

        keyboard_listener.stop()
        recorder.stop_recording()
        recorder.save_recording(filename)

        is_recording.set(False)
        status_var.set(f"录制完成: 已保存到 {os.path.basename(filename)}")

        messagebox.showinfo("完成", f"录制已保存到 {filename}", parent=root)

    except Exception as e:
        is_recording.set(False)
        status_var.set(f"录制失败: {str(e)}")
        messagebox.showerror("录制失败", str(e), parent=root)
    finally:
        recorder_instance = None

def start_playback(recording_file, playback_times, root, is_playing, status_var, config):
    global player_instance
    import os
    from tkinter import messagebox

    if is_playing.get():
        messagebox.showwarning("提示", "正在回放中，请先停止回放")
        return

    filename = recording_file.get()
    times = playback_times.get()

    if not filename or not os.path.exists(filename):
        messagebox.showerror("错误", f"文件不存在: {filename}")
        status_var.set(f"错误: 文件不存在")
        return

    try:
        is_playing.set(True)
        status_var.set(f"回放中... (共 {times} 次)")

        # 更新配置
        config['last_script'] = filename
        save_config(config)

        player = MousePlayer(filename)
        player_instance = player

        # 创建进度更新线程
        def update_progress():
            event_count = len(player.data['events'])
            current_event = 0
            while is_playing.get() and current_event < event_count:
                time.sleep(0.1)
                # 估算进度
                progress = int((current_event / event_count) * 100)
                status_var.set(f"回放中... 进度: {progress}%")
                current_event += 1

        progress_thread = threading.Thread(target=update_progress, daemon=True)
        progress_thread.start()

        # 回放
        player.play(times=times, blocking=True)

        is_playing.set(False)
        status_var.set("回放完成")

        messagebox.showinfo("完成", "回放完成", parent=root)

    except Exception as e:
        is_playing.set(False)
        status_var.set(f"回放失败: {str(e)}")
        messagebox.showerror("回放失败", str(e), parent=root)
    finally:
        player_instance = None

def select_window_title(root, window_title_var):
    # 弹出窗口选择对话框，支持点击选择
    import win32gui
    from tkinter import messagebox

    titles = []
    def enum_windows(hwnd, _):
        title = win32gui.GetWindowText(hwnd)
        if win32gui.IsWindowVisible(hwnd) and title:
            titles.append(title)
    win32gui.EnumWindows(enum_windows, None)

    if not titles:
        messagebox.showerror("错误", "未找到可用窗口", parent=root)
        return

    # 创建选择窗口
    dialog = tk.Toplevel(root)
    dialog.title("选择窗口")
    dialog.geometry("500x400")
    dialog.transient(root)
    dialog.grab_set()

    # 搜索框
    search_frame = tk.Frame(dialog)
    search_frame.pack(fill="x", padx=10, pady=10)

    tk.Label(search_frame, text="搜索窗口:", font=("微软雅黑", 10)).pack(anchor="w")
    search_var = tk.StringVar()
    search_entry = tk.Entry(search_frame, textvariable=search_var, font=("微软雅黑", 10))
    search_entry.pack(fill="x", pady=(5, 0))
    search_entry.focus()

    # 窗口列表
    list_frame = tk.Frame(dialog)
    list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    scrollbar = tk.Scrollbar(list_frame)
    scrollbar.pack(side="right", fill="y")

    listbox = tk.Listbox(list_frame, font=("微软雅黑", 10), yscrollcommand=scrollbar.set)
    listbox.pack(fill="both", expand=True)
    scrollbar.config(command=listbox.yview)

    # 提示标签
    info_label = tk.Label(dialog, text=f"找到 {len(titles)} 个窗口", font=("微软雅黑", 9), fg="#666")
    info_label.pack(pady=(0, 5))

    # 填充列表
    def refresh_list(filter_text=""):
        listbox.delete(0, tk.END)
        filtered = [t for t in titles if filter_text.lower() in t.lower()]
        for title in filtered:
            listbox.insert(tk.END, title)
        info_label.config(text=f"找到 {len(filtered)} 个窗口")

    refresh_list()

    # 搜索功能
    def on_search(event=None):
        refresh_list(search_var.get())

    search_entry.bind("<KeyRelease>", on_search)

    # 双击选择
    def on_select(event=None):
        selection = listbox.curselection()
        if selection:
            window_title_var.set(listbox.get(selection[0]))
            dialog.destroy()

    listbox.bind("<Double-Button-1>", on_select)

    # 回车键选择
    def on_return(event=None):
        on_select()

    listbox.bind("<Return>", on_return)

    # 确认按钮
    btn_frame = tk.Frame(dialog)
    btn_frame.pack(fill="x", padx=10, pady=(0, 10))

    tk.Button(btn_frame, text="确定", command=on_select, width=10).pack(side="right", padx=5)
    tk.Button(btn_frame, text="取消", command=dialog.destroy, width=10).pack(side="right", padx=5)

    # 居中显示
    dialog.update_idletasks()
    x = root.winfo_x() + (root.winfo_width() // 2) - (dialog.winfo_width() // 2)
    y = root.winfo_y() + (root.winfo_height() // 2) - (dialog.winfo_height() // 2)
    dialog.geometry(f"+{x}+{y}")

# 使用示例
if __name__ == "__main__":
    import sys
    import os
    import traceback

    try:
        # 替换原有命令行和交互式菜单入口为 GUI
        gui_main()
    except Exception as e:
        # 如果出现错误，显示错误信息
        error_msg = f"程序启动失败: {str(e)}\n\n{traceback.format_exc()}"

        # 尝试显示错误对话框
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()  # 隐藏主窗口
            messagebox.showerror("启动错误", error_msg)
            root.destroy()
        except:
            pass

        # 打印到控制台
        print(error_msg)
        input("按回车键退出...")  # 防止窗口立即关闭