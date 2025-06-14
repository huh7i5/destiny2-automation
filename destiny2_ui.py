# destiny2_ui.py

import sys
import os
import threading
import time
import ctypes
import queue

import cv2
import numpy as np
from PIL import ImageGrab

import win32gui
import win32con
import win32api

import keyboard
import tkinter as tk
from tkinter import scrolledtext

# —— 资源根路径 —— 
if getattr(sys, "frozen", False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

WINDOW_TITLE = "Destiny 2"

# —— 模板定义 —— 
TEMPLATES = {
    'start':  (os.path.join(base_path, 'start.png'),  0.8),
    'a':      (os.path.join(base_path, 'a.png'),      0.8),
    'return': (os.path.join(base_path, 'return.png'), 0.8),
}

# —— SendInput 结构定义 —— 
PUL = ctypes.POINTER(ctypes.c_ulong)
class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long), ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", PUL),
    ]

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong), ("dwExtraInfo", PUL),
    ]

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", ctypes.c_ulong),
        ("wParamL", ctypes.c_short),
        ("wParamH", ctypes.c_ushort),
    ]

class INPUT(ctypes.Structure):
    class _I(ctypes.Union):
        _fields_ = [
            ("mi", MOUSEINPUT), ("ki", KEYBDINPUT),
            ("hi", HARDWAREINPUT)
        ]
    _anonymous_ = ("_i",)
    _fields_ = [
        ("type", ctypes.c_ulong),
        ("_i", _I),
    ]

def send_input(inputs):
    arr = (INPUT * len(inputs))(*inputs)
    ctypes.windll.user32.SendInput(len(inputs), arr, ctypes.sizeof(INPUT))

# —— 窗口居中函数 —— 
def center_window(hwnd, width=None, height=None):
    """
    如果 width/height 提供，则先调整窗口大小到 width×height；
    然后把窗口移动到屏幕正中央（仅对窗口化/无边框窗口有效）。
    """
    screen_w = win32api.GetSystemMetrics(0)
    screen_h = win32api.GetSystemMetrics(1)
    # 获取当前窗口外框尺寸
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    cur_w = right - left
    cur_h = bottom - top
    # 如果指定了目标大小，使用它
    tgt_w = width  if width  is not None else cur_w
    tgt_h = height if height is not None else cur_h
    new_x = (screen_w - tgt_w) // 2
    new_y = (screen_h - tgt_h) // 2
    win32gui.MoveWindow(hwnd, new_x, new_y, tgt_w, tgt_h, True)
    time.sleep(0.2)

# —— 鼠标操作 —— 
def mouse_move(x, y):
    w = win32api.GetSystemMetrics(0); h = win32api.GetSystemMetrics(1)
    ax = int(x * 65535 / (w - 1)); ay = int(y * 65535 / (h - 1))
    mi = MOUSEINPUT(dx=ax, dy=ay, mouseData=0,
                    dwFlags=win32con.MOUSEEVENTF_MOVE|win32con.MOUSEEVENTF_ABSOLUTE,
                    time=0, dwExtraInfo=None)
    send_input([INPUT(type=0, mi=mi)])

def mouse_click():
    down = MOUSEINPUT(dx=0, dy=0, mouseData=0,
                      dwFlags=win32con.MOUSEEVENTF_LEFTDOWN,
                      time=0, dwExtraInfo=None)
    up   = MOUSEINPUT(dx=0, dy=0, mouseData=0,
                      dwFlags=win32con.MOUSEEVENTF_LEFTUP,
                      time=0, dwExtraInfo=None)
    send_input([INPUT(type=0, mi=down), INPUT(type=0, mi=up)])

def click_at(x, y):
    mouse_move(x, y); time.sleep(0.02); mouse_click()

# —— 键盘操作 —— 
def press_key(key, duration):
    vk_map = {'W':0x57, 'D':0x44, 'O':0x4F}
    vk = vk_map[key.upper()]
    down = KEYBDINPUT(wVk=vk, wScan=0, dwFlags=0, time=0, dwExtraInfo=None)
    up   = KEYBDINPUT(wVk=vk, wScan=0,
                      dwFlags=win32con.KEYEVENTF_KEYUP,
                      time=0, dwExtraInfo=None)
    send_input([INPUT(type=1, ki=down)]); time.sleep(duration); send_input([INPUT(type=1, ki=up)])

# —— 获取游戏窗口句柄并准备 —— 
def prepare_window():
    hwnd = win32gui.FindWindow(None, WINDOW_TITLE)
    if not hwnd:
        raise Exception(f"未找到窗口: {WINDOW_TITLE}")
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.2)
    # 先调整到 1920×1080 并居中
    center_window(hwnd, 1920, 1080)
    return hwnd

# —— 根据 hwnd 计算客户区在屏幕上的矩形 —— 
def get_client_rect(hwnd):
    l, t, r, b = win32gui.GetClientRect(hwnd)
    sl, st = win32gui.ClientToScreen(hwnd, (l, t))
    return (sl, st, sl + (r - l), st + (b - t))

# —— 模板匹配 —— 
def find_template(path, rect, thresh=0.8):
    img = cv2.cvtColor(np.array(ImageGrab.grab(bbox=rect)), cv2.COLOR_BGR2GRAY)
    tpl = cv2.imread(path, 0)
    h, w = tpl.shape
    res = cv2.matchTemplate(img, tpl, cv2.TM_CCOEFF_NORMED)
    _, mv, _, ml = cv2.minMaxLoc(res)
    if mv >= thresh:
        return (rect[0] + ml[0] + w//2, rect[1] + ml[1] + h//2)
    return None

# —— 全局状态 —— 
stop_event = threading.Event()
worker_thread = None
log_queue = queue.Queue()
current_count = 0
max_iterations = None

def queue_log(msg):
    log_queue.put(msg)

# —— 自动化主循环 —— 
def automation_loop():
    global current_count
    current_count = 0

    while not stop_event.is_set():
        # 1) 准备窗口：恢复、调整到 1920×1080 并居中
        try:
            hwnd = prepare_window()
            rect = get_client_rect(hwnd)
            queue_log(f"窗口已调整并居中，客户区坐标: {rect}")
        except Exception as e:
            queue_log(str(e)); queue_log("脚本已停止。"); return

        # 2) 搜 start.png 并点击
        queue_log("搜索 start.png …")
        p = None
        while not stop_event.is_set() and p is None:
            p = find_template(TEMPLATES['start'][0], rect, TEMPLATES['start'][1])
            time.sleep(0.5)
        if stop_event.is_set(): break
        queue_log(f"点击 start: {p}")
        click_at(*p)

        # 3) 等 a.png → W15s, D4s, W5s
        queue_log("等待 a.png …")
        p = None
        while not stop_event.is_set() and p is None:
            p = find_template(TEMPLATES['a'][0], rect, TEMPLATES['a'][1])
            time.sleep(1)
        if stop_event.is_set(): break
        queue_log("检测到 a.png，执行按键序列")
        press_key('W', 15); press_key('D', 4); press_key('W', 5)

        # 4) 等 return.png → O8s
        queue_log("等待 return.png …")
        p = None
        while not stop_event.is_set() and p is None:
            p = find_template(TEMPLATES['return'][0], rect, TEMPLATES['return'][1])
            time.sleep(1)
        if stop_event.is_set(): break
        queue_log("检测到 return.png，长按 O")
        press_key('O', 8)

        # 5) 记录并延时
        current_count += 1
        queue_log(f"本轮完成，第 {current_count} 次，1s 后重试。")
        time.sleep(1)

    queue_log("脚本已停止。")

# —— 启动/停止 & UI —— 
def start_automation():
    global worker_thread, max_iterations, current_count
    if worker_thread and worker_thread.is_alive(): return
    try:
        mi = int(iter_entry.get()); max_iterations = mi if mi>0 else None
    except:
        max_iterations = None
    current_count = 0
    status_label.config(text="运行中")
    count_label.config(text=f"已执行: {current_count}/{max_iterations or '∞'}")
    log_text.config(state='normal'); log_text.delete('1.0', tk.END); log_text.config(state='disabled')
    stop_event.clear()
    worker_thread = threading.Thread(target=automation_loop, daemon=True)
    worker_thread.start()

def stop_automation():
    stop_event.set(); status_label.config(text="已停止")

def process_log_queue():
    try:
        while True:
            msg = log_queue.get_nowait()
            log_text.config(state='normal'); log_text.insert(tk.END, msg+"\n"); log_text.see("end"); log_text.config(state='disabled')
    except queue.Empty:
        pass
    root.after(100, process_log_queue)

def update_count_label():
    count_label.config(text=f"已执行: {current_count}/{max_iterations or '∞'}")
    root.after(500, update_count_label)

keyboard.add_hotkey('f11', start_automation)
keyboard.add_hotkey('end', stop_automation)

root = tk.Tk()
root.title("Destiny2 自动化")
root.geometry("400x360")

frame = tk.Frame(root)
tk.Label(frame, text="最大执行次数 (0=无限):").pack(side="left")
iter_entry = tk.Entry(frame, width=5); iter_entry.insert(0,"0"); iter_entry.pack(side="left",padx=5)
tk.Button(frame, text="开始 (F11)", command=start_automation).pack(side="left",padx=5)
tk.Button(frame, text="停止 (End)",  command=stop_automation).pack(side="left",padx=5)
frame.pack(pady=8)

status_label = tk.Label(root, text="已停止", font=("Arial",12)); status_label.pack()
count_label  = tk.Label(root, text="已执行: 0/∞", font=("Arial",12)); count_label.pack()
log_text     = scrolledtext.ScrolledText(root, state='disabled', width=48, height=12); log_text.pack(pady=8)

root.after(100, process_log_queue)
root.after(500, update_count_label)
root.mainloop()
