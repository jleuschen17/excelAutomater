
import pandas as pd
import yfinance as yf
import datetime as dt
import xlsxwriter
import math
import statistics
from matplotlib import pyplot as plt
import numpy as np



fileName = input("Enter a name for the excel file: ")
def generate_df(ticker, start_date='2021-1-1', volatilityDays=50, calendarYear=365):
    df = yf.download(ticker, start_date, end = '2022-06-30')
    df = df.drop(["Open", "High", "Low", "Volume", "Close"], axis=1)
    df = df.reset_index()
    df["Px Change"] = None
    df["Volatility"] = None
    closes = df["Adj Close"].tolist()
    for i in range(1, len(df)):
        df.iat[i, 2] =round((math.log(df.iloc[i, 1] / df.iloc[i-1, 1])),
                            4)
        if i >= volatilityDays:
            df.iat[i, 3] = round((statistics.stdev(closes[i-50:i]) * math.sqrt(calendarYear)), 4) / 100
    df['Adj Close'] = df['Adj Close'].apply(lambda x: round(x, 2))
    volatilities = df["Volatility"][50:].tolist()
    meanVol = statistics.mean(volatilities)
    prices = df["Adj Close"].tolist()
    meanPrice = statistics.mean(prices)
    factor = meanPrice / meanVol
    df["Adj Vol"] = df["Volatility"] * factor
    for i in range(50):
        df.iat[i, 4] = None
    return df



stocks = ["AMAM", "HCWB", "JANX", "HOWL", "IKNA", "BOLT", "SNSE",
          "CGEM", "SBTX", "ONCR", "NRIX", "ALXO", "ITOS", "AAPL"]

def generate_dfs(stocks, calendarYear, volatilityDays, start_date):
    dfs = {}
    for stock in stocks:
        df = generate_df(stock, start_date, volatilityDays, calendarYear)
        dfs[stock] = df
    return dfs

print("Gathering Data")
dfs = generate_dfs(stocks, 365, 50, dt.date(2015, 1, 1))




mainWorkbook = xlsxwriter.Workbook(f'{fileName}.xlsx')

def writeToWorksheet(df, wb, ticker, calenderYear=365,
                      volatilityDays=50):
    ws = wb.add_worksheet(ticker)
    ws.set_column(0, 0, 12)
    ws.set_column(1, 1, 10)
    ws.set_column(2, 4, 10)
    menu_format = wb.add_format({'bg_color' : '#004481',
                                 'font_color' : 'white',
                                 'font_name' : 'Arial',
                                 'font_size' : 10})
    menu_responses = wb.add_format({'font_name' : 'Arial',
                                    'font_size' : 10,
                                    'align' : 'right'})
    olive_background = wb.add_format({'font_name' : 'Arial',
                                      'font_size' : 10,
                                      'bg_color' : '#cabc96'})
    olive_background_pct = wb.add_format({'font_name' : 'Arial',
                                          'font_size' : 10,
                                          'bg_color' : '#cabc96',
                                          'num_format' : '0.00%'})
    grey_background = wb.add_format({'font_name' : 'Arial',
                                     'font_size' : 10,
                                     'bg_color' : '#eff1ef'})
    grey_background_num = wb.add_format({'font_name' : 'Arial',
                                         'font_size' : 10,
                                         'bg_color' : '#eff1ef',
                                         'num_format' : '0.00%'})
    blue_background = wb.add_format({'font_name' : 'Arial',
                                     'font_size' : 10,
                                     'bg_color' : '#baeafc'})
    menuWords = ["Source", "Ticker", "Calendar Year", "Volatility Days",
                 "Start Date", "Close", "Current", "Average", "Maximum",
                                                              "Minimum"]
    for i in range(len(menuWords)):
        ws.write(i, 0, menuWords[i], menu_format)
    ws.write(0, 1, "Yahoo", menu_responses)
    ws.write(1, 1, ticker, menu_responses)
    ws.write(2, 1, calenderYear, menu_responses)
    ws.write(3, 1, volatilityDays, menu_responses)
    ws.write(4, 1, df.iloc[0][0].strftime("%m/%d/%Y"), menu_responses)
    ws.write(10, 0, "Date", grey_background)
    ws.write(10, 1, "Adj Close", grey_background)
    ws.write(10, 2, "Px Change", grey_background)
    ws.write(10, 3, "Volatility", grey_background)
    ws.write(5, 1, df.iloc[len(df)-1][1], olive_background)
    ws.write(10, 4, "Adj Vol", grey_background)
    volatilities = (df["Volatility"].tolist())[volatilityDays:]
    if len(df) > volatilityDays:
        ws.write(6, 1, df.iloc[len(df)-1][3], olive_background_pct)
        ws.write(7, 1, statistics.mean(volatilities),
                 olive_background_pct)
        ws.write(8, 1, max(volatilities), olive_background_pct)
        ws.write(9, 1, min(volatilities), olive_background_pct)
    for i in range(len(df)):
        try:
            ws.write(i+11, 0, df.iloc[i][0].strftime("%m/%d/%Y"),
                     blue_background)
            ws.write(i+11, 1, df.iloc[i][1], blue_background)
            ws.write(i+11, 2, df.iloc[i][2], grey_background_num)
            ws.write(i+11, 3, df.iloc[i][3], grey_background_num)
            ws.write(i+11, 4, df.iloc[i][4], grey_background_num)
        except:
            pass
    chart = wb.add_chart({'type' : 'line'})
    numCells = len(df)
    chart.add_series({'values' : f'={ticker}!$B$12:$B${str(numCells + 11)}', 'name' : f'={ticker}!$B$11'})
    chart.add_series({'values' : f'={ticker}!$E${volatilityDays}:$E${str(numCells + 11)}', 'name' : f'={ticker}!$D$11'})
    ws.insert_chart('F11', chart)




print("Compiling Excel Sheets")
for stock in stocks:
    writeToWorksheet(dfs[stock], mainWorkbook, stock)

mainWorkbook.close()

print(f"File saved as {fileName}.xlsx")
