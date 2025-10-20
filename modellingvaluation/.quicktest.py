import pandas as pd
from pathlib import Path

INPUT_FILE = "MEDC_FS_Consolidated.xlsx"
if not Path(INPUT_FILE).exists():
    print("File not found:", INPUT_FILE)
else:
    x = pd.ExcelFile(INPUT_FILE)
    print("Sheets in workbook:", x.sheet_names)
