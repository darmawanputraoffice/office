import os
import pyautogui as ag
import pynput as pn

# Setting up working environment
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Press spacebar to get coord
def on_press(key):
    if key == pn.keyboard.Key.space:
        x, y = ag.position()
        print(x, y)
with pn.keyboard.Listener(on_press=on_press) as listener:
    listener.join()