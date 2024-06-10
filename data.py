import yfinance as yf
import pandas as pd
import webbrowser

config_file = open("config.txt", "r")
config = config_file.readlines()

# Get ticker from user, put it to all caps, defaults to MSFT
ticker = input("Ticker: ").strip().upper() or config[0].strip().upper()
stock = yf.Ticker(ticker)

# Get income statement, balance sheet, and cashflow
income = stock.income_stmt
balance = stock.balance_sheet
cashflow = stock.cashflow

# Get path from user, make file called data.xlsx
path_in = config[1].strip()
path = path_in + "data.xlsx"

open_browser = input("Open websites? Y or N: ").strip().upper() or config[2].strip().upper()

# Make ExcelWriter
writer = pd.ExcelWriter(path)

# Put the three tables from before together, and transpose
result = pd.concat([income, balance, cashflow])
resultT = result.transpose()

# List of stats of interest
stat_list = ["Total Revenue", "Cost Of Revenue", "Other Income Expense", "Reconciled Depreciation", "Capital Expenditure", "Working Capital", "Cash Cash Equivalents And Short Term Investments"]
expense_list = ["Research And Development", "Selling General And Administration", "Other Operating Expenses"]
debt_list = ["Current Debt And Capital Lease Obligation", "Long Term Debt And Capital Lease Obligation"]

# Clean up result with the stat_list, drop empty years, then transpose again
clean = resultT[resultT.columns.intersection(stat_list)]
clean = clean[clean["Total Revenue"].notna()]
cleanT = clean.transpose()
cleanT = cleanT.sort_index(axis=1, ascending=True)

other = resultT[resultT.columns.intersection(expense_list)]
otherT = other.transpose()
otherT = otherT.sort_index(axis=1, ascending=True)

debt = resultT[resultT.columns.intersection(debt_list)]
debtT = debt.transpose()
debtT = debtT.sort_index(axis=1, ascending=True)

# Write to Excel
cleanT.to_excel(writer, sheet_name = "data")
sheet = writer.sheets["data"]

otherT.to_excel(writer, sheet_name = "data2")
sheet2 = writer.sheets["data2"]

debtT.to_excel(writer, sheet_name = "data3")
sheet3 = writer.sheets["data3"]

# Make everything in a currency format
fmt_currency = writer.book.add_format({"num_format" : "$#,##0" ,"bold" :False})
sheet.set_column("A:A", 30)
sheet.set_column("B:E", 20, fmt_currency)

sheet2.set_column("A:A", 30)
sheet2.set_column("B:E", 20, fmt_currency)

sheet3.set_column("A:A", 30)
sheet3.set_column("B:E", 20, fmt_currency)

# Close ExcelWriter
writer.close()

analysis_link = "https://finance.yahoo.com/quote/" + ticker + "/analysis/"
WACC_NASDAQ = "https://finbox.com/NASDAQGS:" + ticker + "/models/wacc/"
WACC_NYSE = "https://finbox.com/NYSE:" + ticker + "/models/wacc/"


if open_browser == "Y":
    webbrowser.open(analysis_link)
    if len(ticker) >= 4:
        webbrowser.open(WACC_NASDAQ)
    else:
        webbrowser.open(WACC_NYSE)