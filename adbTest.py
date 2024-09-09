import subprocess
import time
import os
import datetime
import sys
import threading
import keyboard
import cv2
import numpy as np
from functools import lru_cache
import requests

"""
    雷電模擬器:平板版(1280*720)
"""
# 初始化全局變量
keep_running = True  # 控制程序運行狀態

def send_line_notification(message):
    """
    發送LINE通知
    """
    line_notify_token = "E7KfPtMFFERzq8FmaAVBexs7ZQmc22DDCThv0bQwoOi"
    line_notify_api = "https://notify-api.line.me/api/notify"
    headers = {"Authorization": f"Bearer {line_notify_token}"}
    data = {"message": message}
    requests.post(line_notify_api, headers=headers, data=data)

def setup_adb():
    """
    設置 ADB 連接
    """
    # 啟動 ADB 服務器
    subprocess.run("adb start-server", shell=True)
    
    # 獲取已連接設備列表
    devices = run_adb_command("devices")
    if "device" not in devices:
        print("未檢測到已連接的設備，請確保模擬器已啟動並已連接。")
        sys.exit(1)
    
    print("ADB 連接已建立。")
    
# def setup_logging():
#     log_file = "ADB程序日誌.txt"
#     sys.stdout = open(log_file, "w", encoding="utf-8")
#     sys.stderr = sys.stdout

def stop_program():
    global keep_running
    keep_running = False
    print("\n檢測到鍵盤輸入，程序將停止運行。")

def stop_program_on_keypress():
    keyboard.add_hotkey('esc', stop_program)
    print("已設置按下 'esc' 鍵以停止程序。")

def run_adb_command(command):
    """
    執行ADB命令
    """
    full_command = f"adb {command}"
    result = subprocess.run(full_command, shell=True, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"ADB命令執行失敗: {full_command}")
        print(f"錯誤信息: {result.stderr}")
    return result.stdout.strip()

def tap(x, y):
    """
    在指定坐標點擊
    """
    run_adb_command(f"shell input tap {x} {y}")
    print(f"點擊坐標: ({x}, {y})")

def swipe(x1, y1, x2, y2, duration=500):
    """
    從一個坐標滑動到另一個坐標
    """
    run_adb_command(f"shell input swipe {x1} {y1} {x2} {y2} {duration}")
    print(f"滑動: 從 ({x1}, {y1}) 到 ({x2}, {y2})")

def check_image(image_path):
    """
    使用 OpenCV 檢查屏幕上是否存在指定圖像
    """
    try:
        # 使用 ADB 命令捕獲屏幕並直接讀取到內存
        result = subprocess.run("adb exec-out screencap -p", shell=True, capture_output=True)
        if result.returncode != 0:
            print(f"ADB screencap 命令失敗: {result.stderr.decode('utf-8')}")
            return False, None, None

        # 將捕獲的圖像數據轉換為 numpy 數組
        screen_np = np.frombuffer(result.stdout, np.uint8)
        screen = cv2.imdecode(screen_np, cv2.IMREAD_COLOR)

        # 讀取模板圖像
        template = cv2.imread(image_path)
        
        # 使用 OpenCV 進行模板匹配
        result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
        
        if max_val >= 0.8:
            return True, max_loc, template.shape
        else:
            return False, None, None
    except Exception as e:
        print(f"圖像處理過程中發生錯誤: {str(e)}")
        return False, None, None

def find_and_click_image(image_path, max_attempts=100, delay=0.1):
    """
    找到屏幕上的圖像並點擊，如果失敗則重試
    """
    for attempt in range(max_attempts):
        if not keep_running:
            print("程序停止中...")
            return False

        found, location, shape = check_image(image_path)
        if found:
            center_x = location[0] + shape[1] // 2
            center_y = location[1] + shape[0] // 2
            tap(center_x, center_y)
            print(f"點擊圖像: {image_path}")
            return True
        else:
            # print(f"未找到匹配的圖像，重試中... (嘗試 {attempt + 1}/{max_attempts})")
            time.sleep(delay)
    
    print(f"在 {max_attempts} 次嘗試後仍未找到匹配的圖像。")
    return False

def click_images_in_sequence(image_paths, max_attempts=50, delay=0.5):
    """
    依序點擊多張圖片
    """
    for i, image_path in enumerate(image_paths, 1):
        if not keep_running:
            print("程序停止中...")
            return False

        if not os.path.isfile(image_path):
            print(f"文件不存在: {image_path}")
            continue
        
        print(f"正在嘗試點擊第 {i} 張圖片: {image_path}")
        if find_and_click_image(image_path, max_attempts=max_attempts, delay=delay):
            print(f"成功點擊第 {i} 張圖片: {image_path}")
        else:
            print(f"無法點擊第 {i} 張圖片: {image_path}，繼續下一張")
        time.sleep(delay)
    return True

def click_until_next_image(click_coords, next_image_path, max_attempts=50, delay=0.5):
    """
    持續點擊指定坐標，直到能夠檢測到下一張圖片
    """
    for attempt in range(max_attempts):
        if not keep_running:
            print("程序停止中...")
            return False

        tap(click_coords[0], click_coords[1])
        print(f"嘗試 {attempt + 1}/{max_attempts}")
        
        found, _, _= check_image(next_image_path)
        if found:
            print(f"檢測到下一張圖片: {next_image_path}")
            return True
        
        time.sleep(delay)
    
    print(f"在 {max_attempts} 次嘗試後仍未檢測到下一張圖片。")
    return False

def capture_screen():
    result = subprocess.run("adb exec-out screencap -p", shell=True, capture_output=True)
    screen_np = np.frombuffer(result.stdout, np.uint8)
    return cv2.imdecode(screen_np, cv2.IMREAD_COLOR)

def check_image_in_screen(screen, image_path):
    template = cv2.imread(image_path)
    result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, _ = cv2.minMaxLoc(result)
    return max_val >= 0.8

# 主程序
def main():
    
    global keep_running
    
    # 啟動鍵盤監聽
    stop_program_on_keypress()

    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"\n--- 程序開始執行 {current_time} ---\n")
    
    # 定義圖片的路徑
    login = [f"./adb_photo/{i}.png" for i in range(1, 8)]
    take_gift = [f"./adb_photo/take{i}.png" for i in range(0, 3)]
    begin = [f"./adb_photo/begin{i}.png" for i in range(0, 3)]
    beginNext = [f"./adb_photo/beNext{i}.png" for i in range(0, 2)]
    role = [f"./adb_photo/role{i}.png" for i in range(0, 2)]
    dele_path = [f"./adb_photo/dele_{i}.png" for i in range(1, 3)]
    dele_last = [f"./adb_photo/dele_{i}.png" for i in range(3, 5)]
    
    while keep_running:
        setup_adb()
        if find_and_click_image("./adb_photo/0.png"):
            time.sleep(5)
            if click_images_in_sequence(login):
                time.sleep(1)
                if click_until_next_image((1020, 43), "./adb_photo/take0.png"):
                    if click_images_in_sequence(take_gift):
                        if click_images_in_sequence(begin):
                            time.sleep(1)
                            if click_until_next_image((1197, 29), "./adb_photo/skip_close.png"):
                                if find_and_click_image("./adb_photo/skip_close.png"):
                                    """pickup.png為當前卡池"""
                                    if find_and_click_image("./adb_photo/pickup.png"):
                                        if click_images_in_sequence(beginNext):
                                            time.sleep(1)
                                            if click_until_next_image((1197, 29), "./adb_photo/skip_next.png"):
                                                if find_and_click_image("./adb_photo/skip_next.png"):
                                                    if find_and_click_image("./adb_photo/20th_1.png"):
                                                        time.sleep(1)
                                                        if click_until_next_image((1197, 29), "./adb_photo/skip_close.png"):
                                                            if find_and_click_image("./adb_photo/skip_close.png"):
                                                                """genaral.png為普通卡池"""
                                                                if find_and_click_image("./adb_photo/genaral.png"):
                                                                    if click_images_in_sequence(beginNext):
                                                                        time.sleep(1)
                                                                        if click_until_next_image((1197, 29), "./adb_photo/skip_next.png"):
                                                                            if find_and_click_image("./adb_photo/skip_next.png"):
                                                                                if find_and_click_image("./adb_photo/20th_1.png"):
                                                                                    time.sleep(1)
                                                                                    if click_until_next_image((1197, 29), "./adb_photo/skip_close.png"):
                                                                                        if find_and_click_image("./adb_photo/skip_close.png"):
                                                                                            if click_images_in_sequence(role):
                                                                                                print("開始檢查SS\n")

        time.sleep(5)
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
        screen = capture_screen()
    
        found_characters = []
        for image, name in characters.items():
            if check_image_in_screen(screen, f'./adb_photo/{image}'):
                found_characters.append(name)
        
        num = len(found_characters)
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 將結果寫入文件
        with open("ADB首抽結果.txt", "a", encoding="utf-8") as f:
            f.write(f"---{current_time}---\n")
            if found_characters:
                for name in found_characters:
                    f.write(f"{name}\n")
            else:
                f.write("好可悲，一隻都沒有\n")
        
        print(f"找到 {num} 個角色")
        
        if num <= 3:
            if click_until_next_image((50, 14), "./adb_photo/dele_1.png"):
                time.sleep(1)
                if click_images_in_sequence(dele_path):
                    time.sleep(10)
                    if click_images_in_sequence(dele_last):
                        print("刪除賬號\n")
        else:
            notification_message = f"找到 {num} 個角色：\n" + "\n".join(found_characters)
            send_line_notification(notification_message)
            break
        
    print(f"--- 程序執行結束 {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---")
    sys.stdout.close()
    sys.stdout = sys.__stdout__
    sys.stderr = sys.__stderr__

if __name__ == "__main__":
    # 使用 threading 以避免 keyboard 模塊阻塞主線程
    main_thread = threading.Thread(target=main)
    main_thread.start()
    main_thread.join()