import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

import tradinglimit as tl

portofolios = { 
    "case1": {
        "input": {
            "CASHT2": 30000001024,
        },
        "output": {
            "BBCA": 85500002304,
        },
    },
    "case2": {
        "input": {
            "CASHT2": 89999998976,
        },
        "output": {
            "BBCA": 256500006912,
        },
    },
    "case3": {
        "input": {
            "CASHT2": 87594000384,
            "BBCA": {
                "lot": 10000,
                "price": 2430,
            },
        },
        "output": {
            "BBCA": 253913628672,
        },
    },
    "case4": {
        "input": {
            "CASHT2": 87594000384,
            "BBCA": {
                "lot": 10000,
                "price": 2040,
            },
        },
        "output": {
            "BBCA": 253228199957,
        },
    },
    "case5": {
        "input": {
            "CASHT2": 85548900352,
            "BBCA": {
                "lot": 20000,
                "price": 2040,
            },
        },
        "output": {
            "BBCA": 250984964912,
        },
    },
    "case6": {
        "input": {
            "CASHT2": 77368500224,
            "BBCA": {
                "lot": 60000,
                "price": 2040,
            },
        },
        "output": {
            "BBCA": 242012024732,
        }
    },
    "case7": {
        "input": {
            "CASHT2": 73278300160,
            "BBCA": {
                "lot": 80000,
                "price": 2040,
            },
        },
        "output": {
            "BBCA": 230811904726,
        }
    },
    "case8": {
        "input": {
            "CASHT2": 89949872128,
            "ACES": {
                "lot": 10000,
                "price": 50,
            },
        },
        "output": {
            "BBCA": 256431143753,
            "BREN": 0,
        }
    },
    "case9": {
        "input": {
            "CASHT2": 86987489280,
            "ACES": {
                "lot": 10000,
                "price": 50,
            },
            "ADMR": {
                "lot": 30000,
                "price": 985,
            },
        },
        "output": {
            "BBCA": 249901702008,
            "BREN": 0,
        }
    },
    "case10": {
        "input": {
            "CASHT2": 66535493632,
            "ACES": {
                "lot": 10000,
                "price": 50,
            },
            "ADMR": {
                "lot": 30000,
                "price": 985,
            },
            "ASII": {
                "lot": 40199,
                "price": 5075,
            },
        },
        "output": {
            "BBCA": 213582266039,
            "BREN": 0,
        }
    },
}
'''
    "case": {
        "input": {
            "CASHT2": ,
            "BBCA": {
                "lot": ,
                "price": ,
            },
        },
        "output": {
            "BBCA": ,
            "BREN": ,
        }
    },
'''

incorrect = 0
correct = 0

for casenumber, portofolio in portofolios.items():
    for stock_buy, output_value in portofolio["output"].items():
        predict = tl.tradinglimit(account = "FREE", stock_buy = stock_buy, portofolio = portofolio["input"])
        delta = predict - portofolio["output"][stock_buy]
        print("input: " + casenumber + " on " + stock_buy + \
            "\n output: " + str(predict) + \
            "\n output: " + str(portofolio["output"][stock_buy]) + \
            "\n delta : " + str(predict - portofolio["output"][stock_buy]))
        if abs(delta) > 10000:
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            incorrect += 1
        else:
            print("ok")
            correct += 1

print("\n TOTAL X:" + str(incorrect) + \
    "\n TOTAL V:" + str(correct))
