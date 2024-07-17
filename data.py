import yfinance as yf
import pandas as pd
import webbrowser

# Open config file and read each parameter into an array
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

# Ask if user wants to open websites automatically
open_browser = input("Open websites? Y or N: ").strip().upper() or config[2].strip().upper()

# Make ExcelWriter
writer = pd.ExcelWriter(path)

# Put the three tables from before together, and transpose
result = pd.concat([income, balance, cashflow])
resultT = result.transpose()
resultT = resultT[resultT["Total Revenue"].notna()]

# List of stats of interest
stat_list = ["Total Revenue", "Reconciled Depreciation", "Capital Expenditure", "Working Capital", "Cash Cash Equivalents And Short Term Investments", "Cash Cash Equivalents And Federal Funds Sold"]
expense_list = ["Research And Development", "Selling General And Administration", "Other Operating Expenses", "Loss Adjustment Expense", "Occupancy And Equipment", "Other Non Interest Expense", "Professional Expense And Contract Services Expense", "Other Taxes"]
debt_list = ["Current Debt And Capital Lease Obligation", "Long Term Debt And Capital Lease Obligation"]

# Clean up results with the stat_list, drop empty years, then transpose again, sort by year
clean = resultT[resultT.columns.intersection(stat_list)]
cleanT = clean.transpose()
cleanT = cleanT.sort_index(axis=1, ascending=True)

expenses = resultT[resultT.columns.intersection(expense_list)]
expensesT = expenses.transpose()
expensesT = expensesT.sort_index(axis=1, ascending=True)

debt = resultT[resultT.columns.intersection(debt_list)]
debtT = debt.transpose()
debtT = debtT.sort_index(axis=1, ascending=True)

cogs = resultT[resultT.columns.intersection(["Cost Of Revenue"])]
cogsT = cogs.transpose()
cogsT = cogsT.sort_index(axis=1, ascending=True)

other = resultT[resultT.columns.intersection(["Other Income Expense"])]
otherT = other.transpose()
otherT = otherT.sort_index(axis=1, ascending=True)

special = resultT[resultT.columns.intersection(["Special Income Charges", "Other Special Charges"])]
specialT = special.transpose()
specialT = specialT.sort_index(axis=1, ascending=True)

# Write the three separate datasets to Excel sheets
cleanT.to_excel(writer, sheet_name = "main")
sheet1 = writer.sheets["main"]

expensesT.to_excel(writer, sheet_name = "expenses")
sheet2 = writer.sheets["expenses"]

debtT.to_excel(writer, sheet_name = "debt")
sheet3 = writer.sheets["debt"]

cogsT.to_excel(writer, sheet_name = "COGS")
sheet4 = writer.sheets["COGS"]

otherT.to_excel(writer, sheet_name = "other")
sheet4 = writer.sheets["other"]

specialT.to_excel(writer, sheet_name = "special")
sheet5 = writer.sheets["special"]

# Put everything into a currency format in case user wants to read data.xlsx directly
fmt_currency = writer.book.add_format({"num_format" : "$#,##0" ,"bold" :False})
sheet1.set_column("A:A", 30)
sheet1.set_column("B:E", 20, fmt_currency)

sheet2.set_column("A:A", 30)
sheet2.set_column("B:E", 20, fmt_currency)

sheet3.set_column("A:A", 30)
sheet3.set_column("B:E", 20, fmt_currency)

sheet4.set_column("A:A", 30)
sheet4.set_column("B:E", 20, fmt_currency)

sheet5.set_column("A:A", 30)
sheet5.set_column("B:E", 20, fmt_currency)

# Close ExcelWriter
writer.close()

# Concat URLs for websites to be opened
analysis_link = "https://finance.yahoo.com/quote/" + ticker + "/analysis/"
WACC_NASDAQ = "https://finbox.com/NASDAQGS:" + ticker + "/models/wacc/"
WACC_NYSE = "https://finbox.com/NYSE:" + ticker + "/models/wacc/"

# If user wants to open websites, open them
if open_browser == "Y":
    webbrowser.open(analysis_link)
    # Ticker length being 4 or above is usually able to check whether a stock is in NASDAQ or NYSE, other exchanges not yet supported for my finbox links
    if len(ticker) >= 4:
        webbrowser.open(WACC_NASDAQ)
    else:
        webbrowser.open(WACC_NYSE)
        
# To allow for any errors to be read if run like an executable
if config[3].strip().upper() == "Y":
    input("...")