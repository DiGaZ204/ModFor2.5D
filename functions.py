import subprocess
import psutil
import pyautogui
import cv2
import numpy as np
import time
import os
import sys
import keyboard

# 初始化全局變數
keep_running = True  # 控制程式運行狀態

def setup_logging():
    log_file = "程序日誌.txt"
    sys.stdout = open(log_file, "w", encoding="utf-8")
    sys.stderr = sys.stdout

def stop_program():
    global keep_running
    keep_running = False
    print("\n偵測到鍵盤輸入，程式將停止運行。")

def stop_program_on_keypress():
    # 設置按下 'esc' 鍵時停止程式
    keyboard.add_hotkey('esc', stop_program)
    print("已設置按下 'esc' 鍵以停止程式。")

def is_program_running(program_name):
    """
    檢查是否有任何進程與指定的程式名稱匹配
    """
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if program_name.lower() in proc.info['name'].lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False

def open_program(program_path):
    """
    開啟指定的程式
    """
    try:
        subprocess.Popen([program_path])
        print(f"已成功開啟程式: {program_path}")
    except FileNotFoundError:
        print(f"找不到程式: {program_path}")
    except Exception as e:
        print(f"無法開啟程式: {e}")

def find_and_click_icon(icon_path, confidence=0.8, max_attempts=40, delay=0.5):
    """
    找到螢幕上的圖示並點擊，如果失敗則重試
    """
    for attempt in range(max_attempts):
        if not keep_running:
            print("程式停止中...")
            return False

        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        template = cv2.imread(icon_path, cv2.IMREAD_UNCHANGED)
        
        if template is None:
            print(f"无法加载图像: {icon_path}")
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
            icon_height, icon_width = template.shape[:2]
            center_x = max_loc[0] + icon_width // 2
            center_y = max_loc[1] + icon_height // 2
            pyautogui.moveTo(center_x, center_y)
            pyautogui.click()
            print(f"點擊位置: ({center_x}, {center_y})")
            return True
        else:
            print(f"未找到匹配的圖示，重試中... (嘗試 {attempt + 1}/{max_attempts})")
            time.sleep(delay)
    
    print(f"在 {max_attempts} 次嘗試後仍未找到匹配的圖示。")
    return False

def check_file_exists(file_path):
    if os.path.isfile(file_path):
        print(f"文件存在: {file_path}")
    else:
        print(f"文件不存在: {file_path}")

def move_window_to_top_left(window_title):
    try:
        import win32gui
        import win32con

        # 找到目標視窗
        hwnd = win32gui.FindWindow(None, window_title)
        
        if hwnd:
            # 獲取當前視窗位置和大小
            rect = win32gui.GetWindowRect(hwnd)
            x, y, width, height = rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]
            
            # 移動視窗到 (0,0) 位置，保持原有大小
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, width, height, win32con.SWP_SHOWWINDOW)
            print(f"已將 {window_title} 視窗移動到螢幕左上角 (0,0) 位置")
        else:
            print(f"未找到 {window_title} 視窗")
    except ImportError:
        print("無法導入 win32gui 和 win32con 模組，請確保已安裝 pywin32")

def click_images_in_sequence(image_paths, max_attempts=40, delay=0.5):
    """
    依序點擊多張圖片
    """
    for i, image_path in enumerate(image_paths, 1):
        if not keep_running:
            print("程式停止中...")
            return False

        if not os.path.isfile(image_path):
            print(f"文件不存在: {image_path}")
            continue
        
        print(f"正在嘗試點擊第 {i} 張圖片: {image_path}")
        if find_and_click_icon(image_path, max_attempts=max_attempts, delay=delay):
            print(f"成功點擊第 {i} 張圖片: {image_path}")
        else:
            print(f"無法點擊第 {i} 張圖片: {image_path}，繼續下一張")
        time.sleep(delay)
    return True

def click_until_next_image(click_coords, next_image_path, max_attempts=40, delay=0.5):
    """
    持續點擊指定坐標，直到能夠檢測到下一張圖片
    """
    for attempt in range(max_attempts):
        if not keep_running:
            print("程式停止中...")
            return False

        pyautogui.click(click_coords[0], click_coords[1])
        print(f"點擊坐標 {click_coords}，嘗試 {attempt + 1}/{max_attempts}")
        
        if find_icon(next_image_path):
            print(f"檢測到下一張圖片: {next_image_path}")
            return True
        
        time.sleep(delay)
    
    print(f"在 {max_attempts} 次嘗試後仍未檢測到下一張圖片。")
    return False

def find_icon(icon_path, confidence=0.8):
    """
    檢查螢幕上是否存在指定的圖示
    """
    screenshot = pyautogui.screenshot()
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    template = cv2.imread(icon_path, cv2.IMREAD_UNCHANGED)
    
    if template is None:
        print(f"无法加载图像: {icon_path}")
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

    return max_val >= confidence

def switch_to_window(window_title):
    """
    切換到指定標題的視窗
    """
    window = pyautogui.getWindowsWithTitle(window_title)
    if window:
        window[0].activate()
        print(f"已切換到視窗: {window_title}")
        return True
    else:
        print(f"未找到視窗: {window_title}")
        return False
    
def close_chrome():
    os.system("taskkill /f /im chrome.exe")