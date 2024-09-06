import win32gui
import win32ui
import win32con
from ctypes import windll
from PIL import Image
import numpy as np
import cv2
import os

def capture_window_content(window_title):
    hwnd = win32gui.FindWindow(None, window_title)
    if not hwnd:
        print(f"未找到窗口: {window_title}")
        return None

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top

    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()

    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    saveDC.SelectObject(saveBitMap)

    result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)

    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)
    im = Image.frombuffer(
        'RGB',
        (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
        bmpstr, 'raw', 'BGRX', 0, 1)

    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)

    if result == 1:
        # 保存捕獲的窗口畫面
        im.save("captured_window.png")
        print("已保存捕獲的窗口畫面：captured_window.png")
        return np.array(im)
    else:
        print(f"無法捕獲窗口 {window_title} 的內容")
        return None

def find_icon(icon_path, confidence=0.8, window_title="AngelsonStage"):
    screenshot = capture_window_content(window_title)
    if screenshot is None:
        print(f"無法在窗口 {window_title} 中搜索圖標")
        return False

    screenshot = cv2.cvtColor(screenshot, cv2.COLOR_RGB2BGR)
    template = cv2.imread(icon_path, cv2.IMREAD_UNCHANGED)
    
    if template is None:
        print(f"無法加載圖像: {icon_path}")
        return False

    if len(template.shape) == 2:
        screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
        template_gray = template
    elif len(template.shape) == 3 and template.shape[2] == 4:
        screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
        template_gray = cv2.cvtColor(template, cv2.COLOR_BGRA2GRAY)
    elif len(template.shape) == 3 and template.shape[2] == 3:
        screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

    result = cv2.matchTemplate(screenshot_gray, template_gray, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

    if max_val >= confidence:
        # 在截圖上標記找到的圖標位置
        top_left = max_loc
        h, w = template_gray.shape
        bottom_right = (top_left[0] + w, top_left[1] + h)
        cv2.rectangle(screenshot, top_left, bottom_right, (0, 255, 0), 2)
        
        # 保存標記後的截圖
        cv2.imwrite("found_icon.png", screenshot)
        print(f"已保存找到圖標的截圖：found_icon.png")
        
        return True
    else:
        print(f"未找到圖標，最大匹配值：{max_val}")
        return False

# 使用示例
icon_path = './photo2.5/1.png'
window_title = "AngelsonStage"

if os.path.exists(icon_path):
    print(f"圖標文件存在：{icon_path}")
else:
    print(f"圖標文件不存在：{icon_path}")

if find_icon(icon_path, window_title=window_title):
    print("找到圖標")
else:
    print("未找到圖標")
import win32gui

def enum_windows():
    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            windows.append(win32gui.GetWindowText(hwnd))
    windows = []
    win32gui.EnumWindows(callback, windows)
    return windows

print(enum_windows())