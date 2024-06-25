# stock-data-puller
Pulls stock data from Yahoo Finance with the yfinance API to be used in a Discounted Cash Flow

Features:
- [x] Gets ticker from user
- [X] Allow user to edit a config file to set defaults
- [x] Pulls income statement, balance sheet, and cash flow for the company
- [x] Puts into a data.xlsx Excel file to be read by other Excel files (such as one for a DCF)
- [ ] Scrape WACC and Yahoo Finance analyst numbers (currently just opens the websites to be manually grabbed)
	- [X] Otherwise find a way to only open one finbox link (currently approximated with ticker length, can be improved)
- [ ] Allow for more expense names as needed
- [X] Allow for companies with no Cost of Revenue
- [X] Need to figure out how to get "Loss Adjustment Expense" expense (ex: ELV) (fixed in yfinance library)
- [ ] Possibly change the way companies with less than 4 years of Yahoo finances are put into Excel
- [ ] Split off all variables into separate sheets for the best variablility
- [ ] Deal with removing NA values for each one individually
- [ ] Make my own DCF sheet (currently using a proprietary one)

Config file:
With config.txt, you can set defaults and settings as the following:
First line is a default ticker to pull if none is inputted, (default as MSFT)
Second line is a default file location
Third line is a default Y or N for whether to open web links (default Y)
Fourth line is a Y or N to add a pause at the end of execution for debugging purposes (default N)