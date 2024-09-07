import datetime
import threading
from functions import *

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
            'a.png': "天乃里里沙(PICKUP)",
            'b.png': "橘美",
            'c.png': "真由梨",
            'd.png': "雅",
            'e.png': "刃牙子",
            'f.png': "鈴木",
            'g.png': "限定真由梨"
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
            if num < 2:
                if click_images_in_sequence(dele_path):
                    time.sleep(11)
                    find_and_click_icon("./photo2.5/dele_4.png")
                    print("刪除帳號\n")
                    time.sleep(2)  # 等待刪除操作完成
                    return True  # 返回 True 表示需要重新執行
    print(f"\n--- 程序執行結束 {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")
    sys.stdout.close()
    sys.stdout = sys.__stdout__
    sys.stderr = sys.__stderr__
    return False  # 返回 False 表示不需要重新執行

if __name__ == "__main__":
    # 使用 threading 以避免 keyboard 模組阻塞主線程
    def run_main():
        rerun_count = 0
        while True:
            should_rerun = main()
            if should_rerun:
                rerun_count += 1
                if rerun_count % 10 == 0:
                    print("重複執行10次，關閉Chrome")
                    # close_chrome()
                    time.sleep(2)  # 等待Chrome完全關閉
            else:
                break

    main_thread = threading.Thread(target=run_main)
    main_thread.start()
    main_thread.join()