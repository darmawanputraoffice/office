import os
import json
import pandas as pd

os.chdir(os.path.dirname(os.path.abspath(__file__)))

DATAHAIRCUT = pd.read_excel(".ignore/tradinglimitdata.xlsx", sheet_name="Haircut")

CLASS = {
    "A": [0, 25],
    "B": [30, 45],
    "C": [50, 65],
    "D": [70, 75],
    "E": [80, 85],
    "F": [90, 100],
}

HYPARAM = {
    "FREE": {
        "MULTIPLIERSTOCK": {
            "A": 1.85,
            "B": 1.85,
            "C": 1.85,
            "D": 0,
            "E": 0,
            "F": 0,
        },
        "CAPPING": {
            "A": 12.5,
            "B": 6.5,
            "C": 2,
            "D": 0,
            "E": 0,
            "F": 0,
        },
        "EFFECTIVEBUYRATE": {
            "A": 1,
            "B": 1,
            "C": 1,
            "D": 0,
            "E": 0,
            "F": 0,
        },
        "MULTIPLIERCASH": {
            "A": 2.85,
            "B": 2.85,
            "C": 2.85,
            "D": 0,
            "E": 0,
            "F": 0,
        },
    },
    "REGULAR": {
        "MULTIPLIERSTOCK": {
            "A": 1,
            "B": 1,
            "C": 0.7,
            "D": 0.2,
            "E": 0,
            "F": 0,
        },
        "CAPPING": {
            "A": 8,
            "B": 4,
            "C": 2,
            "D": 0,
            "E": 0,
            "F": 0,
        },
        "EFFECTIVEBUYRATE": {
            "A": 1,
            "B": 0.5,
            "C": 0.3,
            "D": 0.3,
            "E": 0.1,
            "F": 0.1,
        },
        "MULTIPLIERCASH": {
            "A": 4,
            "B": 2,
            "C": 1,
            "D": 1,
            "E": 1,
            "F": 1,
        },
    },
    "MARGIN": {
        "MULTIPLIERSTOCK": {
            "A": 1.85,
            "B": 1.85,
            "C": 1.85,
            "D": 0,
            "E": 0,
            "F": 0,
        },
        "CAPPING": {
            "A": 12.5,
            "B": 6.5,
            "C": 2,
            "D": 0,
            "E": 0,
            "F": 0,
        },
        "EFFECTIVEBUYRATE": {
            "A": 1,
            "B": 1,
            "C": 1,
            "D": 0,
            "E": 0,
            "F": 0,
        },
        "MULTIPLIERCASH": {
            "A": 2.85,
            "B": 2.85,
            "C": 2.85,
            "D": 0,
            "E": 0,
            "F": 0,
        },
    },
    "ONLINE": {
        "MULTIPLIERSTOCK": {
            "A": 1,
            "B": 1,
            "C": 0.7,
            "D": 0.2,
            "E": 0,
            "F": 0,
        },
        "CAPPING": {
            "A": 4,
            "B": 2,
            "C": 1,
            "D": 0,
            "E": 0,
            "F": 0,
        },
        "EFFECTIVEBUYRATE": {
            "A": 1,
            "B": 0.5,
            "C": 0.25,
            "D": 0.25,
            "E": 0.07,
            "F": 0.07,
        },
        "MULTIPLIERCASH": {
            "A": 3,
            "B": 2,
            "C": 1,
            "D": 1,
            "E": 1,
            "F": 1,
        },
    },
}

def gethaircut(stock_code):
    haircut = DATAHAIRCUT.loc[DATAHAIRCUT["Kode"] == stock_code, "Haircut"].values[0] / 100
    return haircut

def getclass(stock_code):
    haircut = gethaircut(stock_code)
    for label, [low, high] in CLASS.items():
        if low / 100 <= haircut <= high / 100:
            return label

def tradinglimit(account="FREE", stock_buy="BBCA", portofolio=""):
    stock_class = getclass(stock_buy)
    tl = 0
    for stock_code, stock_info in portofolio.items():
        if stock_code=="CASHT2":
            tl += HYPARAM[account]["MULTIPLIERCASH"][stock_class] * portofolio["CASHT2"]
        else:
            temp_class = getclass(stock_code)
            tl += HYPARAM[account]["MULTIPLIERSTOCK"][temp_class] * \
                min(stock_info["lot"] * 100 * stock_info["price"] * (1 - gethaircut(stock_code)), \
                HYPARAM[account]["CAPPING"][temp_class] * 1000000000)
    tl *= HYPARAM[account]["EFFECTIVEBUYRATE"][stock_class]
    return tl