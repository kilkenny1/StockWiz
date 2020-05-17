try:
    # For Python 3.0 and later
    from urllib.request import urlopen
except ImportError:
    # Fall back to Python 2's urllib2
    from urllib2 import urlopen

import json
from datetime import date
import pandas as pd
from pandas.tseries import offsets
from pandas.tseries.holiday import USFederalHolidayCalendar
import pandas_datareader.data as web
from pandas.tseries.offsets import CustomBusinessDay
import numpy as np
from pandas import json_normalize
import xlwings as xw

def get_jsonparsed_data(url):
  
    response = urlopen(url)
    data = response.read().decode("utf-8")
    df= json_normalize(json.loads(data))
    return df

def get_jsonparsed_data_price(url):
      
    response = urlopen(url)
    data = response.read().decode("utf-8")
    df= json_normalize(json.loads(data)['historical'])
    df['date']=pd.to_datetime(df['date'])
    return df

def get_price(ticker):
    
    url="https://financialmodelingprep.com/api/v3/stock/real-time-price/"
    url=url+ticker
    return get_jsonparsed_data(url)

def get_all_etf():
    url="https://financialmodelingprep.com/api/v3/symbol/available-etfs"
    return get_jsonparsed_data(url)

def get_all_etf_prices():
    url="https://financialmodelingprep.com/api/v3/quotes/etf"
    return get_jsonparsed_data(url)

def get_historical_price(ticker,interval):
    url='https://financialmodelingprep.com/api/v3/historical-chart'+'/'+interval+'/'+ticker
    return get_jsonparsed_data(url)

def get_income_statement(ticker):
    url='https://financialmodelingprep.com/api/v3/financials/income-statement'+'/'+ticker
    return get_jsonparsed_data(url)

def get_historical_price(ticker):
    url='https://financialmodelingprep.com/api/v3/historical-price-full/'+ticker+'?serietype=line'
    return get_jsonparsed_data_price(url)

def get_daily_historical_price(ticker,day:str):
    url='https://financialmodelingprep.com/api/v3/historical-price-full/'+ticker+'?timeseries='+day
    return get_jsonparsed_data_price(url)

def excel_json_parser(data,file_path,worksheet):
    wb=xw.Book(file_path)
    ws=wb.sheets(worksheet)
    ws.cells.clear_contents()
    ws.range('A1').options(index=False).value=data
    print("parse successful")

def get_company_profile(ticker):
    url='https://financialmodelingprep.com/api/v3/company/profile/'+ticker
    return get_jsonparsed_data(url)

def get_company_quote(ticker):
    url='https://financialmodelingprep.com/api/v3/quote/'+ticker
    return get_jsonparsed_data(url)

    
def pct_diff(df,day):
    current=df.loc[0,'close']
    offset=df.loc[(day-1),'close']
    pct_change=(current-offset)/offset
    return pct_change

def get_ytd_price(ticker):
    US_BUSINESS_DAY = CustomBusinessDay(calendar=USFederalHolidayCalendar())
    last_day = date.today() - offsets.YearBegin() - US_BUSINESS_DAY
    last_day=last_day.strftime('%Y-%m-%d')
    url='https://financialmodelingprep.com/api/v3/historical-price-full/'+ticker+'?from='+last_day+'&to='+last_day
    df=get_jsonparsed_data_price(url)
    return df['close'].to_string(index=False)

def get_ytd_pct_change(ticker):
    company_profile=get_company_profile(ticker)
    current=company_profile['profile.price'].to_string(index=False)
    ytd_price=get_ytd_price(ticker)
    pct_change=((float(current)-float(ytd_price))/float(ytd_price))
    return pct_change

def get_peak_alert(ticker):
    company_quote=get_company_quote(ticker)
    current=company_quote['price'][0]
    peak=company_quote['yearHigh'][0]
    if current >= peak :
        return "Strong Sell"
    elif current >= peak*0.95 :
        return "Sell"
    else:
        return "No Alert"
def get_rsi(ticker,days):
    
     window_length=days    
     double_days=str(days*2)
     df=get_daily_historical_price(ticker,double_days)
     
     delta = df['close'].diff()
     delta = delta[1:] 

    # Make the positive gains (up) and negative gains (down) Series
     up=delta.copy()
     down=delta.copy()
     up[up < 0] = 0
     down[down > 0] = 0

    # Calculate the EWMA
     #roll_up1 = up.ewm(span=window_length).mean()
     #roll_down1 = down.abs().ewm(span=window_length).mean()

    # Calculate the RSI based on EWMA
     #RS1 = roll_up1 / roll_down1
     #RSI1 = 100.0 - (100.0 / (1.0 + RS1))

    # Calculate the SMA
     roll_up2 = up.rolling(window_length).mean()
     roll_down2 = down.abs().rolling(window_length).mean()

    # Calculate the RSI based on SMA
     RS2 = roll_up2 / roll_down2
     RSI2 = 100.0 - (100.0 / (1.0 + RS2))
     return(RSI2[-1:].item())

def get_rsi_alert(ticker,day):
    rsi=get_rsi(ticker,day)
    if rsi <=30:
        return 'Strong buy'
    elif rsi <=35:
        return 'Buy'
    elif rsi >=65:
        return 'Sell'
    elif rsi >=70:
        return 'Strong sell'
    else:
        return 'Hold'



def refresh_dashboard(wb,ws,ticker,benchmark):
    #conf 
    wb=xw.Book(wb)
    ws=wb.sheets(ws)

    ##Company profile data
    company_profile=get_company_profile(ticker)
    company_quote=get_company_quote(ticker)
    print('Company profile data for %s get success' %ticker)
    ws.range('B1').value=ticker
    ws.range('A1').value=company_profile['profile.companyName'].to_string(index=False)
    ws.range('B2').value=company_profile['profile.sector'].to_string(index=False)
    ws.range('B3').value=company_profile['profile.industry'].to_string(index=False)
    ws.range('B5').value=company_quote['price'][0]
    ws.range('C5').value=company_quote['change'][0]
    ws.range('D5').value=company_quote['changesPercentage'][0]/100
    ws.range('F1').value=company_quote['earningsAnnouncement'][0][:10]
    print('Company profile for %s updated' %ticker)

    #Benchmark
    benchmark_profile=get_company_profile(benchmark)
    benchmark_price_df=get_daily_historical_price(benchmark,'30')

    print('benchmark %s get success' %benchmark)

    ws.range('A15').value=benchmark_profile['profile.companyName'].to_string(index=False)
    ws.range('B15').value=pct_diff(benchmark_price_df,5)
    ws.range('C15').value=pct_diff(benchmark_price_df,30)
    ws.range('D15').value=get_ytd_pct_change(benchmark)
    print('Benchmark %s updated' %benchmark)

    #Price
    price_df=get_daily_historical_price(ticker,'200')


    print('price data GET success')
    ws.range('B14').value=pct_diff(price_df,5)
    ws.range('C14').value=pct_diff(price_df,30)
    ws.range('D14').value=get_ytd_pct_change(ticker)

    #Technical indicator
    ws.range('A20').value=get_peak_alert(ticker)
    ws.range('B20').value=get_rsi_alert(ticker,14)

    print('Technical indicator of %s updated' %ticker)

    print("Dashboard %s updated successfully" %ws)    


##Conf

excel_json_parser(get_all_etf_prices(),'/Users/Kelvin/Documents/StockAnalyzer.xlsm','ETF')
refresh_dashboard('/Users/Kelvin/Documents/StockAnalyzer.xlsm','INTC','INTC','QQQ')
refresh_dashboard('/Users/Kelvin/Documents/StockAnalyzer.xlsm','NVDA','NVDA','QQQ')
refresh_dashboard('/Users/Kelvin/Documents/StockAnalyzer.xlsm','AAPL','AAPL','QQQ')
refresh_dashboard('/Users/Kelvin/Documents/StockAnalyzer.xlsm','MSFT','MSFT','QQQ')






