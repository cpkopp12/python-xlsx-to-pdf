# when choosing where to start on app without GUI
# just consider flow from input to out put

import pandas as pd
import glob

filepaths = glob.glob("invoices/*.xlsx")

for path in filepaths:
    df = pd.read_excel(path, sheet_name="Sheet 1")
    print(df)