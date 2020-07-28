# SPY-Trade-Analysis
Program utilizes pandas DataReader to pull SPY data from Yahoo Finance and creates an excel file including the analysis needed for short term volatility trading.

Packages needed for program: pandas_datareader, datetime, os, win32com.client

    
1. Creates new folder titled today's date in desired folder
2. User inputs stock ticker criteria
3. Utilizes Yahoo Finance to pull stock data and creates a CSV file
4. Runs macro in Excel to format CSV data
