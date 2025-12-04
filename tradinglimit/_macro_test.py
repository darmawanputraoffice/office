import os
import pyautogui as ag
import pynput as pn
import pytesseract as pt
from PIL import Image, ImageOps
from paddleocr import PaddleOCR

os.chdir(os.path.dirname(os.path.abspath(__file__)))

