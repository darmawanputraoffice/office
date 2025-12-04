import os
import re
import pandas as pd

os.chdir(os.path.dirname(os.path.abspath(__file__)))
rows = []

with open(".ignore/HCPROFINDO_251204.txt", "r", encoding="utf-8") as f:
    for line in f:
        m = re.match(r"\s*\d+\.;([^;]+); *([\d.]+) %; *([\d.]+) %;", line)
        if m:
            share = m.group(1).strip()
            kpei = float(m.group(2))
            panin = float(m.group(3))
            rows.append([share, kpei, panin])
df = pd.DataFrame(rows, columns=["Code", "KPEI", "Panin"])
df.to_excel(".ignore/HCPROFINDO_251204.xlsx", index=False)