import subprocess
import psutil
import pyautogui
import cv2
import numpy as np
import time
import os
import datetime
import sys
import keyboard
import threading  # 新增，用於鍵盤監聽

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

# 將目標程式視窗移動到螢幕左上角 (0,0) 位置
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

# 主程序
def main():
    global keep_running
    setup_logging()
    
    # 啟動鍵盤監聽
    stop_program_on_keypress()

    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"\n--- 程序開始執行 {current_time} ---\n")
    
    program_name = "AngelsonStage.exe"
    program_path = "D:\\dim_2.5\\AngelsonStage.exe"
    window_title = "AngelsonStage"

    # 檢查程式是否正在運行
    if not is_program_running(program_name):
        print(f"{program_name} 未在運行，嘗試開啟它...")
        open_program(program_path)
        time.sleep(5)  # 等待程式啟動
    else:
        print(f"{program_name} 已在運行中。")

    window = pyautogui.getWindowsWithTitle(window_title)
    if window:
        window[0].activate()
    else:
        print(f"未找到窗口: {window_title}")
        exit()
    time.sleep(1)

    # 在程式開始執行前，先將目標視窗移動到左上角
    move_window_to_top_left(window_title)
    time.sleep(1)

    # 定義圖片的路徑
    image_paths_1 = ["./photo2.5/1.png", "./photo2.5/2.png"]
    image_paths_chrome = ["./photo2.5/30.png", "./photo2.5/31.png", "./photo2.5/32.png"]
    image_paths_2 = [f"./photo2.5/{i}.png" for i in range(6, 13)]
    image_paths_3 = [f"./photo2.5/{i}.png" for i in range(13, 17)]
    image_paths_4 = [f"./photo2.5/{i}.png" for i in range(17, 20)]
    image_paths_5 = [f"./photo2.5/{i}.png" for i in range(15, 17)]
    image_paths_6 = [f"./photo2.5/{i}.png" for i in range(18, 20)]
    dele_path = [f"./photo2.5/dele_{i}.png" for i in range(1, 4)]
    time.sleep(1)

    if keep_running:
        if click_until_next_image((1539, 89), "./photo2.5/1.png"):
            if click_images_in_sequence(image_paths_1):
                print("同意後+登入。")
                # 切換到Chrome視窗
                if switch_to_window("SSG LOGIN - Google Chrome"):
                    # 點擊Chrome視窗中的圖片
                    if click_images_in_sequence(image_paths_chrome):
                        print("Chrome視窗中的圖片已成功點擊。")
                        # 切換回AngelsonStage視窗
                        if switch_to_window("AngelsonStage"):
                            time.sleep(1)
                            if find_and_click_icon("./photo2.5/3.png"):
                                print("成功點擊:是。")
                                # 持續點擊 (1539, 89)(禮物位置) 直到檢測到信箱開啟(5)
                                if click_until_next_image((1539, 89), "./photo2.5/6.png"):#信件
                                    print("成功檢測到信件中的(全部領取)。")
                                    # 繼續點擊剩下的圖片
                                    if click_images_in_sequence(image_paths_2):
                                        print("新手十抽中。")
                                        # 持續點SKIP
                                        if click_until_next_image((1805, 73), "./photo2.5/13.png"):#SKIP
                                            time.sleep(1)
                                            print("成功檢測到十抽結束(關閉)。")
                                    if click_images_in_sequence(image_paths_3):
                                        print("進入PICKUP十抽中")
                                        # 持續點SKIP
                                        if click_until_next_image((1805, 73), "./photo2.5/13.png"):#SKIP
                                            time.sleep(1)
                                            print("成功檢測到PICKUP十抽結束(關閉)。")
                                            if find_and_click_icon("./photo2.5/13.png"):
                                                print("關閉，回到卡池")
                                            if click_images_in_sequence(image_paths_5):
                                                print("進入PICKUP第二十抽中")
                                            # 持續點SKIP
                                            if click_until_next_image((1805, 73), "./photo2.5/13.png"):#SKIP
                                                time.sleep(1)
                                                print("成功檢測到PICKUP第二十抽結束(關閉)。")
                                                if find_and_click_icon("./photo2.5/13.png"):
                                                    print("關閉，回到卡池")
                                    if click_images_in_sequence(image_paths_4):
                                        print("進入普池十抽中")
                                        if click_until_next_image((1805, 73), "./photo2.5/13.png"):#SKIP
                                            time.sleep(1)
                                            print("成功檢測到普池十抽結束(關閉)。")
                                            if find_and_click_icon("./photo2.5/13.png"):
                                                print("關閉，回到卡池")
                                            if click_images_in_sequence(image_paths_6):
                                                print("進入普池第二十抽中")
                                            # 持續點SKIP
                                            if click_until_next_image((1805, 73), "./photo2.5/13.png"):#SKIP
                                                time.sleep(1)
                                                print("成功檢測到普池第二十抽結束(關閉)。")
                                                if find_and_click_icon("./photo2.5/13.png"):
                                                    print("關閉，回到卡池")
                                                    time.sleep(2)
                                                    pyautogui.click(1827, 95)
                                                    time.sleep(2)
                                                    if find_and_click_icon("./photo2.5/role.png"):
                                                        time.sleep(2)
        characters = {
            'a.png': "天乃里里沙(T0)    #女主",
            'b.png': "橘美(T0)          #模特",
            'c.png': "真由梨(T0.5)",
            'd.png': "雅",
            'e.png': "刃牙子(T0.5)",
            'f.png': "鈴木",
            'g.png': "限定真由梨(T0)",
            'h.png': "繪理(T1)        #奶媽",
            'i.png': "茜",
            "j.png": "牧野(T0.5)"
        }
        found_characters = []
        for image, name in characters.items():
            if find_icon(f'./photo2.5/{image}'):
                found_characters.append(name)
        num = len(found_characters)
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open("首抽結果.txt", "a", encoding="utf-8") as f:
            f.write(f"---{current_time}---\n")
            if found_characters:
                for name in found_characters:
                    f.write(f"{name}\n")
            else:
                f.write("好可悲，一隻都沒有\n")
            if num < 3:
                if click_images_in_sequence(dele_path):
                    time.sleep(11)
                    find_and_click_icon("./photo2.5/dele_4.png")
                    print("刪除帳號\n")
    print(f"\n--- 程序執行結束 {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")
    sys.stdout.close()
    sys.stdout = sys.__stdout__
    sys.stderr = sys.__stderr__

if __name__ == "__main__":
    # 使用 threading 以避免 keyboard 模組阻塞主線程
    main_thread = threading.Thread(target=main)
    main_thread.start()
    main_thread.join()
