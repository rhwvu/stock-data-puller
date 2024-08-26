# stock-data-puller
Pulls stock data from Yahoo Finance with the yfinance API to be used in a Discounted Cash Flow (DCF)

NOTE: This repo requires the pandas and yfinance packages, and while the pandas version is likely flexible,
	having yfinance 0.2.42 or above is important.

Features:
- [x] Gets ticker from user
- [X] Allow user to edit a config file to set default settings
- [x] Pulls income statement, balance sheet, and cash flow for the ticker's company
- [x] Puts into a data.xlsx Excel file to be read by other Excel files (such as one for a DCF)
- [ ] Scrape WACC (currently opens a good website to be manually copied)
- [X] Gets Yahoo revenue growth numbers (added to yfinance)
- [X] Deals with many individual companies' different expense categorizations
- [ ] Make DCF sheet to go with the project (I'm currently using a proprietary one)

With config.txt, you can set defaults and settings as the following:
- First line is a default ticker to pull if none is inputted, (default as MSFT)
- Second line is a default file location
- Third line is a default Y or N for whether to open web links (default Y)
- Fourth line is a Y or N to add a pause at the end of execution for debugging purposes (default N)