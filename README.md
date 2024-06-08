# stock-data-puller
Pulls stock data from Yahoo Finance with the yfinance API to be used in a Discounted Cash Flow

Features:
- [x] Gets ticker from user
- [X] Allow user to edit a config file to set defaults
- [x] Pulls income statement, balance sheet, and cash flow for the company
- [x] Puts into a data.xlsx Excel file to be read by other Excel files (such as one for a DCF)
- [ ] Probably append data to an existing Excel file instead of making a new one and overwriting it
- [ ] Scrape WACC and Yahoo Finance analyst numbers (currently just opens the websites to be manually grabbed)
- [ ] Allow for more expense names as needed
- [ ] Possibly change the way companies with less than 4 years of Yahoo finances are put into Excel