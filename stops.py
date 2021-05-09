##########################################################################################################################
# imports
import pandas as pd
import yfinance as yf
import numpy as np
import math
from datetime import date
import os
import time
import xlsxwriter
################################################################################################################################
# define function to calculate atr
def calc_atr(ticker): # input is a stock ticker
    """ input a stock ticker, returns a df with ATR """
    ticker_df = yf.download(ticker,
                      start='2020-01-01',                  # one year of data is enough; can adjust later to make this dynamic
                      end=date.today(),
                      progress=False)
    ticker_df['H-L']=abs(ticker_df['High']-ticker_df['Low'])
    ticker_df['H-Cp']=abs(ticker_df.High-ticker_df.Close.shift(1))
    ticker_df['L-Cp']=abs(ticker_df.Low-ticker_df.Close.shift(1))
    ticker_df['TR']=ticker_df[['H-L','H-Cp','L-Cp']].values.max(1)
    ticker_df['ATR']=ticker_df.TR.rolling(window=21).mean()
    ticker_df['ATR2']=(ticker_df.ATR.shift(1)*13+ticker_df.ATR*1)/14
    last=len(ticker_df)
    df_ATR=ticker_df.ATR2
    return df_ATR.iloc[last-1]
##############################################################################################################################
df_stocks=pd.read_excel('stops.xlsx') # import an excel spreadsheet with a list of stocks.

i=0 # initialize counter
print('calculating ATR')
for k in df_stocks.Tickers:             # cycle through each stock in the file and calculate the ATR
    df_stocks.loc[i,'ATR']=calc_atr(k)
    i=i+1
time.sleep(10)                          # pause a bit so that the API does not get fussed

################################################################################################################################
def close(ticker): # input is a stock ticker
    """ input is a stock ticker, output is a dataframe with close prices"""
    ticker_df = yf.download(ticker,
                      start='2020-01-01',                   # one year of data is enough; can adjust later to make this dynamic
                      end=date.today(),
                      progress=False)
    last=len(ticker_df)
    return ticker_df.Close.iloc[last-1]
################################################################################################################################
time.sleep(10)                          # pause a bit so that the API does not get fussed

i=0 # initialize counter
print('getting close prices')
for k in df_stocks.Tickers:  # cycle through each stock in the file and get the close price
    df_stocks.loc[i,'Close']=close(k)
    i=i+1
df_stocks['Stop1.5']=df_stocks.Close-1.5*df_stocks.ATR  # 1.5 ATR stop for short term trading
df_stocks['Stop3']=df_stocks.Close-3*df_stocks.ATR      # 3 ATR stop for short to medium term trading
df_stocks['Stop6']=df_stocks.Close-6*df_stocks.ATR      # 6 ATR stop for longer term investment
df_stocks.set_index('Tickers',inplace=True)
# save the file
print('saving file')
writer = pd.ExcelWriter('stops.xlsx', engine='xlsxwriter')
df_stocks.to_excel(writer,sheet_name='Sheet1')
writer.save()
print('done!!!!!')