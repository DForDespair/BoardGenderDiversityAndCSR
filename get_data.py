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
    tickers=tickers, flds=name_sector_lookup)


# gather firm data for the years 2010-2020
lookup = ['CUR_MKT_CAP', 'BOARD_SIZE', 'INDEPENDENT_DIRECTORS', 'PCT_IND_DIRECTORS_ON_COMP_CMTE',
          'NUMBER_OF_WOMEN_ON_BOARD', 'PCT_WOMEN_ON_BOARD', 'BOARD_AVERAGE_AGE', 'BOARD_MEETING_ATTENDANCE_PCT',
          'BOARD_MEETINGS_PER_YR', 'CEO_DUALITY', 'CEO_AGE', 'TOT_COMP_AW_TO_CEO_&_EQUIV',
          'CHIEF_EXECUTIVE_OFFICER_TENURE', 'CHIEF_EXECUTIVE_OFFICER_AGE', 'FEMALE_CEO_OR_EQUIVALENT',
          'EQY_INST_PCT_SH_OUT', 'ESG_DISCLOSURE_SCORE', 'ENVIRON_DISCLOSURE_SCORE',
          'GOVNCE_DISCLOSURE_SCORE', 'SOCIAL_DISCLOSURE_SCORE']
# weird workaround for lookup as BS data was not loading with regular lookup
bs_lookup = ['BS_TOT_ASSET', 'FNCL_LVRG', 'TOBIN_Q_RATIO',
             'RETURN_ON_ASSET', 'RETURN_COM_EQY', 'SALES_GROWTH', 'EPS_GROWTH', 'PCT_INSIDER_SHARES_OUT', 'SUSTAINALYTICS_RANK', 'SUSTAINALYTICS_GOVERNANCE_PCT'
             'SUSTAINALYTICS_ENVIRONMENT_PCT', 'SUSTAINALYTICS_SOCIAL_PERCENTILE', 'BOARD_AVERAGE_TENURE']
firm_data = blp.bdh(tickers=tickers, flds=lookup,
                    start_date='2010-01-01', end_date='2021-01-01', Per='Y', periodicityAdjustment="ACTUAL")
bs_data = blp.bdh(tickers=tickers, flds=bs_lookup,
                  start_date='2010-01-01', end_date='2021-01-01', Per='Y', periodicityAdjustment="ACTUAL")

# removes multi index format and sets index to date. rename columns
firm_data = firm_data.stack(level=0).reset_index().set_index('level_0')
firm_data.index.rename('Date', inplace=True)
firm_data.rename({'level_1': 'Ticker'}, axis=1, inplace=True)

bs_data = bs_data.stack(level=0).reset_index().set_index('level_0')
bs_data.index.rename('Date', inplace=True)
bs_data.rename({'level_1': 'Ticker'}, axis=1, inplace=True)

# merge dataframes and export to csv
df = firm_data.reset_index().merge(
    bs_data.reset_index(), on=['Date', 'Ticker'])

# merge sector_name table and export to CSV
df2 = pd.merge(df.reset_index(), name_sector_data.reset_index(), left_on='Ticker',
               right_on='index', how='inner')
df2.drop('index_y', axis=1, inplace=True)
df2.drop('index_x', axis=1, inplace=True)
df2.rename({'Date_x': 'Date'}, axis=1, inplace=True)
df2.drop_duplicates(inplace=True)
# %%
df2
# %%
df2.to_csv('./NEWSAMPLE.csv', index=True, header=True)
