import pyautogui
import pyperclip
import openpyxl
import time
import random
import json
import os
from pynput import keyboard

COORDS_FILE = "click_coords.json"
EXCEL_FILE = "input.xlsx"
status = {"running": False, "paused": False, "exit": False}

def random_sleep(a=1, b=8):
    time.sleep(random.uniform(a, b))

def load_coords():
    if os.path.exists(COORDS_FILE):
        with open(COORDS_FILE, 'r') as f:
            return json.load(f)
    return None

def save_coords(coords):
    with open(COORDS_FILE, 'w') as f:
        json.dump(coords, f)

def record_click_positions():
    coords = []
    for i in range(9):
        input(f"[{i+1}/9] 마우스를 {i+1}번째 위치에 두고 Enter를 누르세요.")
        pos = pyautogui.position()
        coords.append(pos)
        print(f"저장됨: {pos}")
    save_coords(coords)
    return coords

def read_excel(filepath):
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    return [[cell.value for cell in row] for row in ws.iter_rows(min_row=2)]

def automation(data, coords, start_row):
    for i, row in enumerate(data[start_row:], start=start_row + 1):
        if status["exit"]:
            print("작업 종료.")
            break

        while status["paused"]:
            print("일시정지 중... p 키로 재개 가능")
            time.sleep(1)

        if not row or not row[0]:
            print(f"{i+1}행 A열이 비어있음. 종료.")
            break

        a, b, c, d, e = row[:5]
        print(f"[{i+1}행] 실행 중...")

        pyautogui.click(coords[0]); random_sleep()
        pyautogui.click(coords[1]); random_sleep()
        pyperclip.copy(str(a)); pyautogui.hotkey('command', 'v'); random_sleep()

        pyautogui.click(coords[2]); random_sleep()
        pyautogui.click(coords[3]); random_sleep()
        pyperclip.copy(str(b)); pyautogui.hotkey('command', 'v'); random_sleep()

        for val in [c, d]:
            pyautogui.press('tab'); 
            pyautogui.press('tab'); random_sleep()
            pyperclip.copy(str(val)); pyautogui.hotkey('command', 'v'); random_sleep()

        for val in [e]:
            pyautogui.press('tab'); random_sleep()
            pyperclip.copy(str(val)); pyautogui.hotkey('command', 'v'); random_sleep()

        for j in range(4, 9):
            pyautogui.click(coords[j]); random_sleep()

    print("자동화 완료!")

def on_key_press(key):
    try:
        if key.char == 's':
            status["running"] = True
            status["paused"] = False
            print("[시작] 자동화 시작됨")

        elif key.char == 'p':
            status["paused"] = not status["paused"]
            print("[일시정지]" if status["paused"] else "[재개]")

        elif key.char == 'q':
            status["exit"] = True
            print("[종료] 자동화 종료 요청됨")
            return False
    except:
        pass

def main():
    coords = load_coords()
    if not coords:
        print("좌표 정보가 없습니다. 클릭 좌표를 새로 등록합니다.")
        coords = record_click_positions()

    data = read_excel(EXCEL_FILE)
    start_row = int(input("시작할 행 번호 (2부터 시작): ")) - 2

    print("s → 시작, p → 일시정지, q → 종료")
    listener = keyboard.Listener(on_press=on_key_press)
    listener.start()

    while not status["running"]:
        time.sleep(0.2)

    automation(data, coords, start_row)

if __name__ == "__main__":
    main()
