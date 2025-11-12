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
#    "case6": { # WEIRD CAUSE LIMIT IS SOMEHOW USED
#        "input": {
#            "CASHT2": 84332478464,
#            "BBCA": {
#                "lot": 25948,
#                "price": 2040,
#            },
#        },
#        "output": {
#            "BBCA": 246784080707,
#        },
#    },
    "case7": {
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
    "case8": {
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
}
'''
    "case": {
        "input": {
            "CASHT2": 0,
            "BBCA": {
                "lot": 0,
                "price": 0,
            },
        },
        "output": {
            "BBCA": 0,
        }
    },
'''

for casenumber, portofolio in portofolios.items():
    for stock_buy, output_value in portofolio["output"].items():
        print("input: " + casenumber + " on " + stock_buy + \
            "\n output: " + str(tl.tradinglimit(account = "FREE", stock_buy = stock_buy, portofolio = portofolio["input"])) + \
            "\n output: " + str(portofolio["output"][stock_buy]))

'''
Finding
1) Capping BBCA di 11,875,000,000
2)
'''