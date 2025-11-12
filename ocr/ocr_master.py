import os
import pytesseract
from PIL import Image
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))

def perform_ocr(image_path):
    try:
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img)
        return text
    except FileNotFoundError:
        return f"Error: Image file '{image_path}' not found."
    except Exception as e:
        return f"Error: {str(e)}"

if __name__ == "__main__":
    for image_file in ["sample", "sample_dewa", "sample_porto"]:
        extracted_text = perform_ocr(image_file + ".png")
        print("Extracted Text:")
        print(extracted_text)
        with open("output_" + image_file + ".txt", "w", encoding="utf-8") as f:
            f.write(extracted_text)