import os
import pytesseract
from PIL import Image, ImageOps

os.chdir(os.path.dirname(os.path.abspath(__file__)))

img = Image.open("sample.png")
#gray = ImageOps.grayscale(img)
text = pytesseract.image_to_string(img, lang="eng")
print(text)
