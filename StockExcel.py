#! python3
"""
Created on Sat Jan 12 18:21:55 2019

Program utilizes pandas DataReader to pull requested stock(s) from Yahoo Finance
with a provided date range.  Creates an excel file of the data

@author: HM
"""

# Imports packages needed for program
import pandas_datareader.data as web
import datetime, os, os.path, win32com.client

wsh = win32com.client.Dispatch("WScript.Shell")

from pandas_datareader._utils import RemoteDataError

# Indicate number of stocks you would like to pull and convert to an Excel file
NumOfStocks = input('Please indicate how many stocks you would like to pull: ')

while not NumOfStocks.isdigit():
    NumOfStocks = input('Please provide an integer for the number of requested stocks: ')
    
# Creates new folder titled today's date in YahooFinance Excel folder (Market Research)
ExcelFolder = 'C:\\Users\\User\\Documents\\Market Research\\Yahoo Finance Excel\\' \
    + datetime.datetime.today().strftime('%Y-%m-%d')

if not os.path.exists(ExcelFolder):
    os.makedirs(ExcelFolder)

# Changes Current Working Directory to newly created folder
os.chdir(ExcelFolder)

DownloadedExcel = []
ValueErrorCount = 0

for i in range(int(NumOfStocks)):
    # Input stock ticker criteria
    ticker = str(input('Please input the ticker of stock #' + str(i + 1) + ': ')).upper()
    while True:
        if os.path.exists('C:\\Users\\User\\Documents\\Market Research\\Yahoo Finance Excel\\' \
                          + datetime.datetime.today().strftime('%Y-%m-%d') + '\\' \
                          + datetime.datetime.today().strftime('%Y-%m-%d') + ' ' + ticker + '.xlsx'):
            ticker = str(input('That stock has already been pulled.  Please input another stock. ')).upper()
        else:
            break

    # Date range criteria
    if i == 0:
        AllData = str(input('Would you like to pull all available data? (Yes/No) ')).upper()
        while not (AllData == 'YES' or AllData == 'NO'):
                AllData = str(input('Please specify Yes or No ')).upper()
        if AllData == 'YES':
            start = '1970-01-01'
            end = datetime.datetime.today().strftime('%Y-%m-%d')
        
        else:
            start = str(input('start date (year-month-day or \'first\'): ')).upper()
            if start == 'FIRST':
                start = '1970-01-01'
            end = str(input('end date (year-month-day or \'today\'): ')).upper()
            if end == 'TODAY':
                end = datetime.datetime.today().strftime('%Y-%m-%d')
    
    # If more than 1 stock, asks if you want to use same dates as previous stock    
    else:
        if ValueErrorCount == 0:
            Answer = str(input('Would you like to use the same dates as the previous stock? (Yes/No) ')).upper()
            while not (Answer == 'YES' or Answer == 'NO'):
                Answer = str(input('Please specify Yes or No ')).upper()
            if Answer == 'YES':
                start == start
                end == end
        else:
            start = str(input('start date (year-month-day or \'first\'): ')).upper()
            if start == 'FIRST':
                start = '1970-01-01'

            end = str(input('end date (year-month-day or \'today\'): ')).upper()
            if end == 'TODAY':
                end = datetime.datetime.today().strftime('%Y-%m-%d')

    # Utilizes input criteria to pull Stock's Yahoo Finance data
    try:
        Stock = web.DataReader(ticker, 'yahoo', start, end)
        Stock.reset_index(inplace=True,drop=False)
        # Creates Excel if no exceptions are raised
        Stock.to_excel(datetime.datetime.today().strftime('%Y-%m-%d') + ' ' + ticker + '.xlsx')
        # Stock.to_csv(ticker + '.csv', sep=' ', encoding='utf-8')  <-- Syntax to create CSV file
        print('\n\tExcel file for ' + ticker + ' has been created.')
        DownloadedExcel.append(ticker)
    
    # Runs macro in Excel to format CSV data (Text to Columns, adjust 
    #     column width, and adds 'Unique ID' in cell A1) 
        if os.path.exists('C:\\Users\\User\\Documents\\Market Research\\Yahoo Finance Excel\\' \
                          + datetime.datetime.today().strftime('%Y-%m-%d') + '\\' \
                          + datetime.datetime.today().strftime('%Y-%m-%d') + ' ' + ticker + '.xlsx'):
            # Tries as if Excel program were open
            try:
                xl = win32com.client.GetActiveObject("Excel.Application.15")
                wb = xl.Workbooks.Open(os.path.abspath(datetime.datetime.today().strftime('%Y-%m-%d') \
                                                       + ' ' + ticker + '.xlsx'))
                xl.Visible = True
                mwb = xl.Workbooks.Open(os.path.abspath('C:\\Users\\User\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB'))
                mwb.Application.Run('PERSONAL.XLSB!YahooFinanceFormat.YahooFinanceFormat')
                wb.Save()
            
            # If Excel program is not open, opens excel, runs macro, and closes file
            except:
                xl = win32com.client.Dispatch("Excel.Application.15")
                wb = xl.Workbooks.Open(os.path.abspath(datetime.datetime.today().strftime('%Y-%m-%d') \
                                                       + ' ' + ticker + '.xlsx'))
                xl.Visible = True
                mwb = xl.Workbooks.Open(os.path.abspath('C:\\Users\\User\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB'))
                mwb.Application.Run('PERSONAL.XLSB!YahooFinanceFormat.YahooFinanceFormat')
                # Saves file and closes Excel connection
                wb.Save()
                wb.Close(True)
                mwb.Close(False)
                xl.Quit()
                del xl

    # Prints error message for either invalid stock ticker or invalid date format
    except RemoteDataError:
        print('ERROR: Stock ticker ' + ticker + ' can not be found on Yahoo Finance')
        
    except ValueError:
        print('ERROR: Please use the correct format (year-month-day) and rerun.')
        ValueErrorCount += 1
    
    except AttributeError:
        continue
        
os.chdir('C:\\Users\\User\\')

print('REMINDER: End Excel process using the Task Manager window.')