import yfinance as yf
import pandas as pd
import webbrowser

# Open config file and read each parameter into a list
config_file = open("config.txt", "r")
config = config_file.readlines()

# Get ticker from user, put it to uppercase, defaults to config[0]
ticker = input("Ticker: ").strip().upper() or config[0].strip().upper()
stock = yf.Ticker(ticker)

# Get income statement, balance sheet, and cashflow
income = stock.income_stmt
balance = stock.balance_sheet
cashflow = stock.cashflow

# Get revenue growth estimates from analysis
growth = stock.revenue_estimate

# Get path from config[1], make file called data.xlsx
path_in = config[1].strip()
path = path_in + "data.xlsx"

# Ask if user wants to open websites automatically, or default to config[2]
open_browser = input("Open websites? Y or N: ").strip().upper() or config[2].strip().upper()

# Make ExcelWriter
writer = pd.ExcelWriter(path)

# Put the three tables together, transpose, drop empty years, and sort by year
result = pd.concat([income, balance, cashflow])
resultT = result.transpose()
resultT = resultT[resultT["Total Revenue"].notna()]
resultT = resultT.sort_index(axis=1, ascending=True)

# Drop empty growth estimates, drop unneeded revenue stats, and transpose
growth = growth[growth["growth"].notna()]
growth = growth["growth"]
growthT = growth.transpose()

# List of stats of interest
stat_list = ["Total Revenue", "Cost Of Revenue", "Reconciled Depreciation", "Capital Expenditure",
    "Working Capital", "Cash Cash Equivalents And Short Term Investments", "Cash Cash Equivalents And Federal Funds Sold",
    "Research And Development", "Selling General And Administration", "Other Operating Expenses",
    "Loss Adjustment Expense", "Occupancy And Equipment", "Other Income Expense", "Other Non Interest Expense",
    "Professional Expense And Contract Services Expense", "Other Taxes", "Current Debt And Capital Lease Obligation",
    "Long Term Debt And Capital Lease Obligation", "Special Income Charges", "Other Special Charges"]

# Clean up results with stat_list and transpose again (while avoiding an unhelpful pandas FutureWarning)
with pd.option_context("future.no_silent_downcasting", True):
    clean = resultT.reindex(columns=stat_list).fillna(0).infer_objects(copy=False)
cleanT = clean.transpose()

# Write the main dataset to a sheet in data.xlsx
cleanT.to_excel(writer, sheet_name = "data")
main_sheet = writer.sheets["data"]

# Write the growth estimate data to a separate sheet in data.xlsx
growthT.to_excel(writer, sheet_name = "growth")
growth_sheet = writer.sheets["growth"]

# Put main data into a currency format in case user wants to read data.xlsx directly
fmt_currency = writer.book.add_format({"num_format" : "$#,##0" ,"bold" :False})
main_sheet.set_column("A:A", 30)
main_sheet.set_column("B:E", 20, fmt_currency)

# Put growth estimate data into a percentage format
fmt_percent = writer.book.add_format({"num_format": "0.0%"})
growth_sheet.set_column("B:B", 10, fmt_percent)


# Close ExcelWriter
writer.close()

# Concatenate URLs for websites to be opened
WACC_NASDAQ = "https://finbox.com/NASDAQGS:" + ticker + "/models/wacc/"
WACC_NYSE = "https://finbox.com/NYSE:" + ticker + "/models/wacc/"

# If user wants to open websites, open them
if open_browser == "Y":
    # Ticker length being 4 or above is usually able to check whether a stock is in NASDAQ or NYSE
    if len(ticker) >= 4:
        webbrowser.open(WACC_NASDAQ)
    else:
        webbrowser.open(WACC_NYSE)
        
# config[3] allows for errors to be read after run complete
if config[3].strip().upper() == "Y":
    input("...")