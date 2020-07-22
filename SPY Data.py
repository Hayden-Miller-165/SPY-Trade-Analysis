#! python3
"""
Created on Sat Jan 12 18:21:55 2019

Program utilizes pandas DataReader to pull SPY data from Yahoo Finance.  
Creates an excel file of the data.

@author: HM
"""

# Imports packages needed for program
import pandas_datareader.data as web
import datetime, os, os.path, win32com.client, sys

answer = ''

while answer.upper() != 'NO' and answer.upper() != 'YES':
    answer = input('Would you like to begin the SPY program?  ')

if answer.upper() == 'NO':
    sys.exit()

print('SPY pull program now initiating....')

wsh = win32com.client.Dispatch("WScript.Shell")
    
# Creates new folder titled today's date in YahooFinance Excel folder (Market Research)
ExcelFolder = 'FILE LOCATION HERE' \
    + datetime.datetime.today().strftime('%Y-%m-%d')

if not os.path.exists(ExcelFolder):
    os.makedirs(ExcelFolder)

# Changes Current Working Directory to newly created folder
os.chdir(ExcelFolder)

DownloadedExcel = []
ValueErrorCount = 0

for i in range(int(1)):
    # Input stock ticker criteria
    ticker = str('SPY')
    while True:
        if os.path.exists('FILE LOCATION HERE' \
                          + datetime.datetime.today().strftime('%Y-%m-%d') + '\\' \
                          + datetime.datetime.today().strftime('%Y-%m-%d') + ' ' + ticker + '.xlsx'):
            sys.exit
        else:
            break

    # Date range criteria
    if i == 0:
            start = '1970-01-01'
            end = datetime.datetime.today().strftime('%Y-%m-%d')

    # Utilizes input criteria to pull Stock's Yahoo Finance data
    try:
        Stock = web.DataReader('AAPL', 'yahoo', start, end)
        Stock.reset_index(inplace=True,drop=False)
        # Creates Excel if no exceptions are raised
        Stock.to_excel(datetime.datetime.today().strftime('%Y-%m-%d') + ' ' + ticker + '.xlsx')
        # Stock.to_csv(ticker + '.csv', sep=' ', encoding='utf-8')  <-- Syntax to create CSV file
        DownloadedExcel.append(ticker)
    
    # Runs macro in Excel to format CSV data (Text to Columns, adjust 
    #     column width, and adds 'Unique ID' in cell A1) 
        if os.path.exists('FILE LOCATION HERE' \
                          + datetime.datetime.today().strftime('%Y-%m-%d') + '\\' \
                          + datetime.datetime.today().strftime('%Y-%m-%d') + ' ' + ticker + '.xlsx'):
            # Tries as if Excel program were open
            try:
                xl = win32com.client.GetActiveObject("Excel.Application.15")
                wb = xl.Workbooks.Open(os.path.abspath(datetime.datetime.today().strftime('%Y-%m-%d') \
                                                       + ' ' + ticker + '.xlsx'))
                xl.Visible = True
                mwb = xl.Workbooks.Open(os.path.abspath('FILE LOCATION HERE'))
                mwb.Application.Run('MACRO LOCATION HERE')
                wb.Save()
            
            # If Excel program is not open, opens excel, runs macro, and closes file
            except:
                xl = win32com.client.Dispatch("Excel.Application.15")
                wb = xl.Workbooks.Open(os.path.abspath(datetime.datetime.today().strftime('%Y-%m-%d') \
                                                       + ' ' + ticker + '.xlsx'))
                xl.Visible = True
                mwb = xl.Workbooks.Open(os.path.abspath('MACRO REPOSITORY LOCATION HERE'))
                mwb.Application.Run('MACRO LOCATION HERE')
                # Saves file and closes Excel connection
                wb.Save()
                wb.Close(True)
                mwb.Close(False)
                xl.Quit()
                del xl
    
    except AttributeError:
        continue