import os
import cv2
import pytesseract

os.chdir(os.path.dirname(os.path.abspath(__file__)))
#print(pytesseract.get_tesseract_version())

img = cv2.imread("sample.png")
gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
text = pytesseract.image_to_string(gray, lang="eng")
print(text)