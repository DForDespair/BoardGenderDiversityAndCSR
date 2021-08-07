import numpy as np
import pandas as pd
from xbbg import blp

path = r"C:\Users\amogh\Desktop\Coding\CSR_Research\data\SAMPLEONE.xlsx"

data = pd.read_excel(path, header=0)
# eliminate excess white space from each column for better parsing
data.columns = [column.strip(' ') for column in data.columns]
# gather unique tickers
tickers = data['Ticker'].unique()

name_sector_lookup = ['Security_Name', 'GICS_Sector_Name']

name_sector_data = blp.bdp(
    tickers='AAPL US Equity', flds=name_sector_lookup)

print(name_sector_data)
