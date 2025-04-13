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
    print("\n[클릭 좌표 등록]")
    print("총 9개의 마우스 위치를 등록합니다. 각 위치에 마우스를 놓고 Enter를 누르세요.\n")
    for i in range(9):
        input(f"[{i+1}/9] 마우스를 원하는 위치에 두고 Enter를 누르세요: ")
        pos = pyautogui.position()
        coords.append(pos)
        print(f"저장됨: {pos}")
    save_coords(coords)
    return coords

def read_excel(filepath):
    if not os.path.exists(filepath):
        print(f"[오류] 엑셀 파일 '{filepath}' 이(가) 존재하지 않습니다.")
        exit(1)
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    return [[cell.value for cell in row] for row in ws.iter_rows(min_row=2)]

def automation(data, coords, start_row):
    for i, row in enumerate(data[start_row:], start=start_row + 1):
        if status["exit"]:
            print("작업 종료 요청됨. 자동화를 종료합니다.")
            break

        while status["paused"]:
            print("[일시정지 중] p 키로 재개 가능")
            time.sleep(1)

        if not row or not row[0]:
            print(f"{i+1}행 A열이 비어있습니다. 자동화를 종료합니다.")
            break

        a, b, c, d, e = row[:5]
        print(f"[{i+1}행] 작업 진행 중...")

        pyautogui.click(coords[0]); random_sleep()
        pyautogui.click(coords[1]); random_sleep()
        pyperclip.copy(str(a)); pyautogui.hotkey('ctrl', 'v'); random_sleep()

        pyautogui.click(coords[2]); random_sleep()
        pyautogui.click(coords[3]); random_sleep()
        pyperclip.copy(str(b)); pyautogui.hotkey('ctrl', 'v'); random_sleep()

        for val in [c, d]:
            pyautogui.press('tab')
            pyautogui.press('tab'); random_sleep()
            pyperclip.copy(str(val)); pyautogui.hotkey('ctrl', 'v'); random_sleep()

        pyautogui.press('tab'); random_sleep()
        pyperclip.copy(str(e)); pyautogui.hotkey('ctrl', 'v'); random_sleep()

        for j in range(4, 9):
            pyautogui.click(coords[j]); random_sleep()

    print("[완료] 자동화가 완료되었습니다.")

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
    print("""
==============================
 자동화 프로그램 (Windows 전용)
==============================

[처음 실행 시 꼭 확인할 사항]
 - input.xlsx 파일을 같은 폴더에 준비하세요.
 - 열 순서는 A열부터 E열까지 아래와 같이 구성되어야 합니다:
   A: 항목1, B: 항목2, C: 항목3, D: 항목4, E: 항목5
 - 엑셀은 반드시 .xlsx 형식이어야 합니다.

[클릭 좌표 등록 안내]
 - 총 9개 좌표를 등록합니다.
 - 이 좌표는 마우스 클릭이 발생할 위치입니다.
 - 반드시 자동화 작업이 이루어질 고정된 크기의 프로그램 창을 열어둔 상태에서 등록해야 합니다.
 - 이 창의 크기나 위치가 바뀌면 자동화가 오작동할 수 있습니다.

[프로그램 실행 위치]
 - 모든 서류 입력을 완료하고 "접수 창"이 열린 상태에서 자동화를 시작해야 합니다.
 - 즉, 자동화 시작 시점은 접수창이 포그라운드로 떠 있는 상태여야 합니다.

[단축키 안내]
  s → 시작
  p → 일시정지 / 재개
  q → 종료

※ 실행 중 키보드로 언제든 제어 가능합니다.
※ 마우스 클릭은 사용자 등록 위치 기준으로 자동 수행됩니다.
==============================
""")
    coords = load_coords()
    if not coords:
        print("[안내] 클릭 좌표 정보가 없습니다. 새로 등록을 시작합니다.")
        coords = record_click_positions()

    data = read_excel(EXCEL_FILE)
    try:
        start_row = int(input("시작할 행 번호를 입력하세요 (예: 2): ")) - 2
    except:
        print("[오류] 유효한 숫자를 입력하지 않아 종료합니다.")
        return

    print("s → 시작, p → 일시정지, q → 종료")
    listener = keyboard.Listener(on_press=on_key_press)
    listener.start()

    while not status["running"]:
        time.sleep(0.2)

    automation(data, coords, start_row)

if __name__ == "__main__":
    main()
