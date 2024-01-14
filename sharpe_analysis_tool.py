import pandas as pd
import numpy as np
import datetime as dt
from dateutil.relativedelta import relativedelta
import yfinance as yf
import matplotlib.pyplot as plt

def main(tickersParam, periodParam: int, rfParam: float, intervalParam: str, modeParam: str):
    """
    Description: Main function that formats the function call parameters and calls other functions

    Parameters:
    - tickersParam: Either a file path to an Excel file containing tickers or a list of tickers
    - periodParam: Number of years of historical data (Integer)
    - rfParam: Risk-free rate for calculating the Sharpe ratio (Float)
    - intervalParam: Time interval for data download (String)
    - modeParam: Data type to be used (String)

    Returns:
    - plot: Matplotlib figure containing the bar plot of the top stocks based on Sharpe ratios
    """
        
    if isinstance(tickersParam, str):
        tickers = pd.read_excel(tickersParam, header= None)
        tickers = tickers[0].tolist()
    elif isinstance(tickersParam, list):
        tickers = tickersParam
    tickers = tickers[:50]
    log_returns = download(tickers, periodParam, intervalParam, modeParam)
    sharpes = calculateSharpe(log_returns, tickers, rfParam)
    plot = plotting(sharpes)

    return(plot)


def download(tickers, periodParam: int, intervalParam: str, modeParam: str):
    """
    Description: Function to download historical stock price data and calculate logarithmic returns.

    Parameters:
    - tickers: List of stock tickers
    - periodParam: Number of years of historical data (Integer)
    - intervalParam: Time interval for data download (String)
    - modeParam: Data type to be used (String)

    Returns:
    - log_returns: DataFrame containing logarithmic returns for each stock
    """
    
    endDate = dt.datetime.now()
    startDate = endDate - relativedelta(years= periodParam)

    data = yf.download(tickers, start= startDate, end= endDate, interval= intervalParam)[modeParam]
    log_returns = np.log(data/data.shift(1))[1:].dropna(axis= 1)

    return(log_returns)


def calculateSharpe(log_returns, tickers, rfParam):
    """
    Description: Function to calculate Sharpe ratios for each stock.

    Parameters:
    - log_returns: DataFrame containing logarithmic returns for each stock
    - tickers: List of stock tickers
    - rfParam: Risk-free rate for calculating the Sharpe ratio (Float)

    Returns:
    - dataframe: DataFrame containing Sharpe ratios for each stock
    """
        
    sharpes = {}
    for ticker in tickers:
        returns = log_returns[ticker]
        stdev = np.std(returns) * np.sqrt(12)
        mean = returns.mean() * 12
        sharpe = (mean - rfParam) / stdev
        sharpes[ticker] = sharpe 

    df = pd.DataFrame([sharpes], columns= None)
    dataframe = df.T
    dataframe = dataframe.rename(columns= {0: 'Sharpe'})
    dataframe = dataframe.sort_values(by= ['Sharpe'], ascending= False)[:10]
    return(dataframe)


def plotting(sharpes):
    """
    Description: Function to create a bar plot of the top stocks based on Sharpe ratios.

    Parameters:
    - sharpes: DataFrame containin Sharpe ratios for each stock

    Returns:
    - fig: Matplotlib figure containing the bar plot
    """
        
    sharpe = sharpes.reset_index()
    fig = plt.figure(figsize=(16, 12))
    title = plt.title(f'Top {len(sharpes)} stocks measured by the Sharpe ratio.', pad=20, fontsize= 15)
    plt.subplots_adjust(top=0.9, bottom=0.1)
    plt.bar(sharpe["index"], sharpe["Sharpe"], width=0.8, color="blue")

    for i, ratio in enumerate(sharpe['Sharpe']):
        if ratio < 0:
            position = 'top'
        else:
            position = 'bottom'
        plt.text(sharpe['index'][i], ratio, str(round(ratio, 2)), horizontalalignment='center', verticalalignment= position, fontsize=10)

    return(fig)

"""
Function Call

Parameters:
- tickersParam: .xlsx file path OR list including ticker names
- periodParam: Number of years of historical data (Integer)
- rfParam: Risk-free rate for calculating the Sharpe ratio in decimal format
- intervalParam: Time interval 1m, 2m, 5m, 15m, 30m, 60m, 90m, 1h, 1d, 5d, 1wk, 1mo, 3mo. Intraday data cannot extend last 60 days
- modeParam: Mode of data (Use always 'Adj Close')
"""

sharpe = main("FILE PATH OR LIST OF TICKERS HERE", 1, 0.025, '1mo', 'Adj Close')
plt.show()
