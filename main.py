import pandas as pd
import glob


filespath = glob.glob("invoices/*.xlsx")

for filepath in filespath:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)