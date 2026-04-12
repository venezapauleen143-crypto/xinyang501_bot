"""
桌面控制腳本 - 供 Claude Code 使用
用法：python desktop_control.py <action> [參數...]

actions:
  screenshot                    截圖並儲存到桌面
  click <x> <y>                 點擊
  double_click <x> <y>          雙擊
  right_click <x> <y>           右鍵點擊
  move <x> <y>                  移動滑鼠
  type <文字>                   輸入文字
  press <按鍵>                  按下按鍵
  open <程式路徑或名稱>         開啟程式
  scroll <up|down> [格數]       滾動
  pos                           取得目前滑鼠座標
"""

import sys
import io
import subprocess
import pyautogui
from pathlib import Path
from datetime import datetime

pyautogui.FAILSAFE = True
SCREENSHOT_DIR = Path.home() / "Desktop"

def screenshot():
    img = pyautogui.screenshot()
    filename = SCREENSHOT_DIR / f"screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    img.save(filename)
    print(f"截圖已儲存：{filename}")
    return str(filename)

def click(x, y):
    pyautogui.click(int(x), int(y))
    print(f"已點擊 ({x}, {y})")

def double_click(x, y):
    pyautogui.doubleClick(int(x), int(y))
    print(f"已雙擊 ({x}, {y})")

def right_click(x, y):
    pyautogui.rightClick(int(x), int(y))
    print(f"已右鍵點擊 ({x}, {y})")

def move(x, y):
    pyautogui.moveTo(int(x), int(y), duration=0.3)
    print(f"滑鼠已移到 ({x}, {y})")

def type_text(text):
    pyautogui.write(text, interval=0.05)
    print(f"已輸入：{text}")

def press_key(key):
    pyautogui.press(key)
    print(f"已按下：{key}")

def open_app(app):
    subprocess.Popen(app, shell=True)
    print(f"已開啟：{app}")

def scroll(direction, amount=3):
    amount = int(amount)
    pyautogui.scroll(amount if direction == "up" else -amount)
    print(f"已向{direction}滾動 {amount} 格")

def pos():
    x, y = pyautogui.position()
    print(f"目前滑鼠位置：({x}, {y})")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    action = sys.argv[1]
    args = sys.argv[2:]

    actions = {
        "screenshot": lambda: screenshot(),
        "click": lambda: click(*args),
        "double_click": lambda: double_click(*args),
        "right_click": lambda: right_click(*args),
        "move": lambda: move(*args),
        "type": lambda: type_text(" ".join(args)),
        "press": lambda: press_key(args[0]),
        "open": lambda: open_app(" ".join(args)),
        "scroll": lambda: scroll(*args),
        "pos": lambda: pos(),
    }

    if action not in actions:
        print(f"未知動作：{action}")
        print(__doc__)
        sys.exit(1)

    try:
        actions[action]()
    except Exception as e:
        print(f"執行失敗：{e}")
        sys.exit(1)
