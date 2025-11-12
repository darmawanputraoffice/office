import os
import easyocr

os.chdir(os.path.dirname(os.path.abspath(__file__)))

reader = easyocr.Reader(['en'])
result = reader.readtext('sample.png')

for box, text, conf in result:
    print(text)