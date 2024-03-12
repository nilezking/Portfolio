import pandas as pd
import numpy as np
import yfinance as yf
import statsmodels.api as sm


def dataRetrieval(tickersParam, indiceParam: str, startParam: str, endParam: str, intervalParam: str, modeParam: str, writeParam: str, pathParam: str):
    """
    Description: Function that downloads market data from Yahoo Finance and calculates returns for every stock and index

    Parameters:
    - tickersParam: Excel file path or list containing tickers
    - indiceParam: Index symbol for market data
    - startParam: Start date for data retrieval
    - endParam: End date for data retrieval
    - intervalParam: Time interval for data
    - modeParam: Mode of data
    - writeParam: Selection between Excel writer and print on terminal
    - pathParam: Path and file name for Excel writer

    Returns:
    - betas: DataFrame with regression results
    """

    if isinstance(tickersParam, str):
        tickers = pd.read_excel(tickersParam, header= None)
        tickers = tickers[0].tolist()
    elif isinstance(tickersParam, list):
        tickers = tickersParam
    
    stock_prices = yf.download(tickers, start= startParam, end= endParam, interval= intervalParam)[modeParam]
    stock_returns = np.log(stock_prices/stock_prices.shift(1))

    market_price = yf.download(indiceParam, start= startParam, end= endParam, interval= intervalParam)[modeParam]
    market_return = np.log(market_price/market_price.shift(1))

    returns = pd.concat([stock_returns, market_return], axis=1)
    returns = returns[1:].dropna(axis= 1)
    nobs = returns.shape[0]

    tickers = list(returns.columns.values)[:-1]
    reg_analyses = regression(returns, tickers)
    betas = compilation(reg_analyses, nobs)

    if writeParam == 'E':
        dataExport(betas, pathParam)
    elif writeParam == 'T':
        return(betas)
        

def regression(returns, tickers):
    """
    Description: Function that runs multiple subsequent regressions and saves results (stock betas) in a dictionary

    Parameters:
    - returns: DataFrame with stock and market returns
    - tickers: List of stock tickers for analysis

    Returns:
    - models: Dictionary with regression models for each stock
    """

    models = {}
    for ticker in tickers:
        X = returns['Adj Close']
        y = returns[ticker]
        X = sm.add_constant(X)
        model = sm.OLS(y, X)
        result = model.fit()
        models[ticker] = result
    return(models)

def compilation(models, nobs):
    """
    Description: Function that compiles regression models, R-squared values and P-values in a DataFrame

    Parameters:
    - models: Dictionary with regression models for each stock
    - nobs: Number of observations used in regression model (Integer)

    Returns:
    - df: DataFrame with compiled regression results
    """

    ticker_list = []
    results_list = []
    rsquared = []
    pvalue = []
    observations = []

    for ticker, result in models.items():
        ticker_list.append(ticker)
        results_list.append(round(result.params["Adj Close"], 3))
        rsquared.append(round(result.rsquared, 3))
        pvalue.append(format(result.pvalues["Adj Close"], 'f'))
        observations.append(nobs)

    complete_data = {'Ticker': ticker_list, 'Beta': results_list, 'R-squared': rsquared, 'P-value': pvalue, 'Obs.': observations}
    df = pd.DataFrame(complete_data)
    return(df)

def dataExport(betas, pathParam):
    """
    Description: Function that exports data into Excel file

    Parameters:
    - betas: DataFrame with regression results (betas)
    - pathParam: Path and file name for Excel writer

    Returns:
    - None
    """

    betas.to_excel(pathParam, index= False)


"""
Function Call

Parameters:
- tickersParam: .xlsx file path OR list including ticker names
- indiceParam: Index symbol for market data
- startParam: Start date for data retrieval in 'yyyy-mm-dd' format
- endParam: End date for data retrieval in 'yyyy-mm-dd' format
- intervalParam: Time interval 1m, 2m, 5m, 15m, 30m, 60m, 90m, 1h, 1d, 5d, 1wk, 1mo, 3mo. Intraday data cannot extend last 60 days
- modeParam: Mode of data (Use always 'Adj Close')
- writeParam: Selection between Excel writer and print on terminal ('E' = Excel, 'T' = Terminal)
"""

beta = dataRetrieval('FILE PATH OR LIST OF TICKERS', '^OMXHGI', '2019-11-01', '2023-12-01', '1mo', 'Adj Close', 'T', 'FILE NAME AND PATH HERE')
print(beta)
