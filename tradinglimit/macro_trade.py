import os
import time
import json
import random as rnd
import pygetwindow as gw
import pyautogui as ag
import dotenv

# Setting up working environment
os.chdir(os.path.dirname(os.path.abspath(__file__)))
dotenv.load_dotenv()

# Preparing Data and Variable
PIN = os.getenv("PIN")
with open("macro_data.json", "r") as f:
    MASTERDATA = json.load(f)
portofolio_load = {}
for stock in MASTERDATA["stock_trade"]:
    portofolio_load[stock] = 0

def transact(
    deal : str = "",
    code : str = "", 
    lot : int = 0,
    client : str = ""
):
    """Buy / Sell Stock"""
    # Open buy / sell menu
    time.sleep(0.3)
    if deal == "buy":
        ag.hotkey("F2")
    elif deal == "sell":
        ag.hotkey("F4")
    else:
        print("Please input deal type")

    # Enter stock code
    if code != "":
        time.sleep(0.3)
        ag.typewrite(code)
    else:
        print("Please input stock code")
    time.sleep(0.3)
    ag.press("tab")
    time.sleep(0.3)
    ag.press("tab")

    # Enter lot
    if lot > 0:
        time.sleep(0.3)
        ag.typewrite(str(lot))
    else:
        print("Please input lot")
    time.sleep(0.3)
    ag.press("tab")
    
    # Enter client
    if client != "":
        time.sleep(0.3)
        ag.typewrite(str(client)) # 88888 / 7126 / 7131
    else:
        print("Please input client")
    time.sleep(0.3)
    ag.press("tab")

def confirm(value : bool = 0):
    """Confirm to buy / sell stock"""
    time.sleep(0.3)
    if value == 0:
        ag.press("esc")
    else:
        ag.press("enter")
        time.sleep(0.3)
        ag.press("enter")

# Switch to Proclick
time.sleep(0.3)
proclick = gw.getWindowsWithTitle("ProCLICK - Profindo Online Trading")[0]
screenshot_size = (proclick.left, proclick.top, proclick.width, proclick.height)
if proclick.isMinimized:
    proclick.restore()
proclick.activate()

# Login PIN Trading
time.sleep(0.3)
ag.hotkey("ctrl", "l")
time.sleep(0.3)
ag.typewrite(PIN)
time.sleep(0.3)
ag.hotkey("enter")

while all(load < 300000 for load in portofolio_load.values()):
    # Make portofolio cases
    stock_choices = [stock for stock, load in portofolio_load.items() if load < 300000]
    stock_choice = rnd.choice(stock_choices)
    lot_choice = rnd.randint(1, 5) * 10000

    transact(
        deal = "sell",
        code = stock_choice,
        lot = lot_choice,
        client = 88888
    )
    confirm(value = 0)
    transact(
        deal = "buy",
        code = stock_choice, 
        lot = lot_choice,
        client = 7126
    )
    confirm(value = 0)
    portofolio_load[stock_choice] += lot_choice 

    # Checking trading limit
    ag.screenshot(region=screenshot_size).save(f".ignore/porto_{MASTERDATA["case_number_trade"]}.png")
    for index, stock in enumerate(MASTERDATA["stock_check"]):
        transact(
            deal = "buy",
            code = stock,
            lot = 1,
            client = 7126
        )
        ag.screenshot(region=screenshot_size).save(f".ignore/porto_{MASTERDATA["case_number_trade"]}_{index + 1}_{stock}.png")
        confirm(value = 0)  

    MASTERDATA["case_number_trade"] += 1
    with open("macro_data.json", "w") as f:
        json.dump(MASTERDATA, f, indent=4)

# Logout PIN Trading
time.sleep(0.3)
ag.hotkey("ctrl", "o")