import pandas
import pandas as pd
import numpy_financial as np
from datetime import date
import os
from openpyxl.workbook import Workbook

pd.set_option('display.max_colwidth', False)
pd.set_option('display.max_rows', False)
pandas.set_option( "display.expand_frame_repr", False)

interest = 0.04
years = 30
payments_year = 12
mortgage = 400000
start_date = (date(2021, 1, 1))

# Monthly Payment
pmt = -1 * np.pmt(interest / payments_year, years * payments_year, mortgage)

# Inteest Payment
ipmt = -1 * np.ipmt(interest / payments_year, 1, years * payments_year, mortgage)

# Principal Payment
ppmt = -1 * np.ppmt(interest / payments_year, 1, years * payments_year, mortgage)

rng = pd.date_range(start_date, periods=years * payments_year, freq="MS")
freq = "MS"
rng.name = "Payment Date"

df = pd.DataFrame(index=rng, columns=["Payment", "Principal Paid", "Interest Paid", "Ending Balance"], dtype="float")
df.reset_index(inplace=True)
df.index += 1
df.index.name = "Period"


df["Payment"] = -1 * np.pmt(interest / payments_year, years * payments_year, mortgage)
df["Principal Paid"] = -1 * np.ppmt(interest / payments_year, df.index, years * payments_year, mortgage)
df["Interest Paid"] = -1 * np.ipmt(interest / payments_year, df.index, years * payments_year, mortgage)
df["Ending Balance"] = 0

df.loc[1, "Ending Balance"] = mortgage - df.loc[1, "Principal Paid"]
df = df.round(2)

for period in range(2, len(df) + 1):
    previous_balance = df.loc[period-1, "Ending Balance"]
    principal_paid = df.loc[period, "Principal Paid"]

    if previous_balance == 0:
        df.loc[period, ["Payment", "Principal Paid", "Interest Paid", "Ending Balance"]] == 0
        continue
    elif principal_paid <= principal_paid:
        df.loc[period, "Ending Balance"] = previous_balance - principal_paid

df.to_excel (r"D:\\03. Python\\Simple Mortgage Calculator With Python and Excel\\Pandas_Mortgage_Calculator.xlsx", index = False, header=True)