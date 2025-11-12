import os
import pyautogui
import time
import pygetwindow as gw
from PIL import ImageGrab

def switch_to_notepad_and_automate(text_to_type, click_x=None, click_y=None, screenshot_path="notepad_screenshot.png"):
    try:
        current_window = gw.getActiveWindow()
        current_window_title = current_window.title if current_window else None
    except:
        current_window_title = None
    notepad_windows = gw.getWindowsWithTitle('Notepad')
    if not notepad_windows:
        print("Notepad is not open!")
        return False
    notepad = notepad_windows[0]
    if notepad.isMinimized:
        notepad.restore()
    notepad.activate()
    time.sleep(1)
    if click_x is not None and click_y is not None:
        pyautogui.click(click_x, click_y)
        time.sleep(0.2)
    pyautogui.write(text_to_type, interval=0.05)
    time.sleep(0.2)
    left, top, width, height = notepad.left, notepad.top, notepad.width, notepad.height
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    screenshot.save(screenshot_path)
    print(f"Screenshot saved to: {screenshot_path}")
    time.sleep(0.2)
    if current_window_title:
        try:
            prev_windows = gw.getWindowsWithTitle(current_window_title)
            if prev_windows:
                prev_windows[0].activate()
                print(f"Switched back to: {current_window_title}")
        except:
            print("Could not switch back to previous window")
    return True

def find_cursor_position():
    print("Move your mouse to desired position. Press Ctrl+C to stop.")
    try:
        while True:
            x, y = pyautogui.position()
            print(f"X: {x}, Y: {y}", end='\r')
            time.sleep(0.1)
    except KeyboardInterrupt:
        print(f"\nFinal position - X: {x}, Y: {y}")
        return x, y

if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    switch_to_notepad_and_automate("Hello from Python automation!")