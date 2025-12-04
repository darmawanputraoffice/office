import os
import json
import pyautogui as ag
import pynput as pn
import pytesseract as pt
from PIL import Image, ImageOps
from paddleocr import PaddleOCR

# Setting up working environment
os.chdir(os.path.dirname(os.path.abspath(__file__)))
ocr = PaddleOCR(lang='en')

# Preparing Data and Variable
value_types = ["casht2", "stock_code", "stock_lot", "stock_price"]
coords = {
    "casht2"      : (2250,  360, 2879,  410),
    "stock_code"  : (   5,  530,  160, 1565),
    "stock_lot"   : ( 305,  530,  430, 1565),
    "stock_price" : ( 425,  530,  550, 1565),
}
tradinglimit_coord = (1120,  860, 1360,  920)
temp_case_input = {}
temp_case_output = {}
temp_anomaly = {}
temp_images = {}
temp_texts = {}
temp_scores = {}
with open("macro_data.json", "r") as f:
    MASTERDATA = json.load(f)

for case_number in range(MASTERDATA["case_number_record"], MASTERDATA["case_number_trade"] + 1):
    image_porto = Image.open(f".ignore/porto_{case_number}.png")
    for value_type in value_types:
        # Cropping images
        temp_images[value_type] = ImageOps.invert(image_porto.crop(coords[value_type]))
        temp_images[value_type].save(f".ignore/temp_{value_type}.png")

        # Predicting portofolio
        ocr_data = ocr.predict(f".ignore/temp_{value_type}.png")
        temp_texts[value_type] = ocr_data[0]["rec_texts"]
        temp_scores[value_type] = ocr_data[0]["rec_scores"]

    # Building portofolio
    temp_case_input["CASHT2"] = temp_texts["casht2"][0]
    for code, lot, price in zip(temp_texts["stock_code"], temp_texts["stock_lot"], temp_texts["stock_price"]):
        temp_case_input[code] = {
            "lot" : lot,
            "price": price
        }
    
    # Checking portofolio anomaly
    if not(len(temp_texts["stock_code"]) == len(temp_texts["stock_lot"]) == len(temp_texts["stock_price"])):
        temp_anomaly["missing"] == True
    for scores in temp_scores.values():
        for score in scores:
            if score < 0.99:
                temp_anomaly["unconfident"] = True
                print(score)
    
    # Finding Trading Limit
    files = [
        file for file in os.listdir(".ignore/")
        if file.startswith(f"porto_{case_number}_") and file.endswith(".png")
    ]
    for file in files:
        image_tradinglimit = Image.open(f".ignore/{file}")
        temp_image = ImageOps.invert(image_tradinglimit.crop(tradinglimit_coord))
        temp_image.save(f".ignore/temp.png")
        ocr_data = ocr.predict(f".ignore/temp.png")
        temp_case_output[file[-8:-4]] = ocr_data[0]["rec_texts"][0]

    # Recording data to .json
    MASTERDATA["testcase"][f"case_{case_number}"] = {
        "input" : temp_case_input,
        "output" : temp_case_output
    }
    MASTERDATA["anomaly"][f"case_{case_number}"] = temp_anomaly
    MASTERDATA["case_number_record"] += 1
    with open("macro_data.json", "w") as f:
        json.dump(MASTERDATA, f, indent=4)

    temp_case_input.clear()
    temp_case_output.clear()
    temp_anomaly.clear()
    temp_images.clear()
    temp_texts.clear()
    temp_scores.clear()