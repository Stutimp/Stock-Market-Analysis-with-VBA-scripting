# Stock Market Analysis using VBA Scripting

## Overview
This project involves creating a VBA script to analyze stock market data across multiple quarters. The script will loop through all stocks for each quarter and output key information, including ticker symbols, quarterly changes, percentage changes, and total stock volume. The project includes additional functionality to identify stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume. The script is designed to run on every worksheet (quarter) at once and includes conditional formatting to highlight changes.

### Features
- **Data Retrieval**: The script loops through stock data for each quarter and retrieves the following values:

Ticker symbol
Volume of stock
Opening price
Closing price

- **Column Creation**: The script creates the following columns for each worksheet:

Ticker symbol
Total stock volume
Quarterly change ($)
Percent change

- **Conditional Formatting: Applied to:**

Quarterly change column (positive change in green, negative change in red)
Percent change column (positive change in green, negative change in red)

- **Calculated Values:**

Greatest percentage increase
Greatest percentage decrease
Greatest total volume

- **Multi-Sheet Processing:** The script runs on all worksheets within the workbook, processing each quarter's data.


-**Conclusion:**
This VBA script automates the analysis of stock market data, making it easier to extract meaningful insights across multiple quarters. It leverages VBA's power to handle repetitive tasks efficiently, providing a robust solution for stock data analysis.

