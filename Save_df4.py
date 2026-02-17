#!/usr/bin/env python
# coding: utf-8

# In[8]:


import os
import time
import requests
import pandas as pd
import numpy as np
import yfinance as yf
import pytz
from datetime import datetime
from pandas.tseries.offsets import BMonthEnd
from dateutil.relativedelta import relativedelta

from sklearn.linear_model import LinearRegression
import scipy.stats as stats
import bs4 as bs

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
import xlsxwriter
import plotly
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns

# ===========================
# PATH CONFIG (GITHUB SAFE)
# ===========================
BASE_PATH = "./data"
OUTPUT_PATH = os.path.join(BASE_PATH, "Comp")
COMP_PATH = os.path.join(BASE_PATH, "CD")
DIVIDEND_PATH = os.path.join(BASE_PATH, "Dividend")
TICKERS_FILE = "./tickers.txt"

API_KEY = '5VJ5OHNNB582TWUU'

STATEMENTS = {
    'INCOME_STATEMENT': 'Income Statement',
    'BALANCE_SHEET': 'Balance Sheet',
    'CASH_FLOW': 'Cash Flow Statement'
}

os.makedirs(OUTPUT_PATH, exist_ok=True)
os.makedirs(COMP_PATH, exist_ok=True)
os.makedirs(DIVIDEND_PATH, exist_ok=True)

# ===========================
# Helper functions
# ===========================
def get_financial_data(symbol, function, retries=5, sleep_time=20):
    url = "https://www.alphavantage.co/query"
    params = {"function": function, "symbol": symbol, "apikey": API_KEY}

    for attempt in range(1, retries + 1):
        response = requests.get(url, params=params)

        if response.status_code != 200:
            time.sleep(sleep_time)
            continue

        data = response.json()

        if "Note" in data or "Information" in data:
            time.sleep(sleep_time)
            continue

        if "quarterlyReports" in data:
            return data["quarterlyReports"][:20]

        time.sleep(sleep_time)

    return []

def create_dataframe(reports, statement_name):
    df = pd.DataFrame()
    for report in reports:
        fiscal_date = pd.to_datetime(report.get('fiscalDateEnding')).strftime('%Y-%m-%d')
        df_quarter = pd.DataFrame.from_dict(report, orient='index')
        df_quarter.columns = [fiscal_date]
        df_quarter.index.name = statement_name
        df = pd.concat([df, df_quarter], axis=1)
    return df

def clean_dynamic_df(df):
    df.columns = [str(c).split(".")[0] for c in df.columns]
    df = df.groupby(df.columns, axis=1).first()

    df.reset_index(inplace=True)
    df.rename(columns={'index': 'Line Item'}, inplace=True)

    if 'Line Item' in df.columns:
        df = df[df['Line Item'] != 'fiscalDateEnding']

    data = {
        'Line Item': df.iloc[:, 0].tolist(),
        **df.iloc[:, 1:].sort_index(axis=1, ascending=False).to_dict(orient='list')
    }

    df_dynamic = pd.DataFrame(data)
    df_dynamic.set_index('Line Item', inplace=True)
    return df_dynamic

def get_company_overview(symbol, retries=5, sleep_time=12):
    url = 'https://www.alphavantage.co/query'
    params = {'function': 'OVERVIEW', 'symbol': symbol, 'apikey': API_KEY}

    for attempt in range(1, retries + 1):
        response = requests.get(url, params=params)

        if response.status_code != 200:
            time.sleep(sleep_time)
            continue

        data = response.json()

        if "Note" in data or "Information" in data:
            time.sleep(sleep_time)
            continue

        if data:
            df = pd.DataFrame.from_dict(data, orient='index', columns=['Value'])
            df.index.name = 'Attribute'
            return df

        time.sleep(sleep_time)

    return None

def process_ticker(ticker):
    combined_df = pd.DataFrame()

    for func, name in STATEMENTS.items():
        reports = get_financial_data(ticker, func)
        if reports:
            df_section = create_dataframe(reports, name)
            combined_df = pd.concat([combined_df, df_section], axis=1)

    if not combined_df.empty:
        df_clean = clean_dynamic_df(combined_df)
        output_file = os.path.join(OUTPUT_PATH, f"df_{ticker}.xlsx")
        df_clean.to_excel(output_file, index=True)

    company_df = get_company_overview(ticker)
    if company_df is not None:
        out_file = os.path.join(COMP_PATH, f"CD_{ticker}.xlsx")
        company_df.to_excel(out_file, index=True)

def export_dividends(ticker):
    timezone = 'America/New_York'
    today = pd.Timestamp.now(tz=pytz.timezone(timezone)).normalize()

    series = yf.Ticker(ticker).dividends

    if not series.empty:
        series = series.sort_index(ascending=False)
        dfdi = series.to_frame(name='Dividend').reset_index().rename(columns={'index': 'Date'})
        dfdi['Date'] = pd.to_datetime(dfdi['Date']).dt.tz_localize(None)
        data = dfdi[['Date', 'Dividend']].copy()
    else:
        first_date = BMonthEnd().rollback(today)
        dates = []
        for i in range(20):
            candidate = first_date - pd.DateOffset(months=3 * i)
            candidate_bme = BMonthEnd().rollback(candidate)
            dates.append(candidate_bme.normalize())
        data = pd.DataFrame({'Date': dates, 'Dividend': 1})

    data['Date'] = pd.to_datetime(data['Date']).dt.strftime('%Y-%m-%d')
    data = data.sort_values('Date', ascending=False).reset_index(drop=True)

    out_file = os.path.join(DIVIDEND_PATH, f"Div_{ticker}.xlsx")
    data.to_excel(out_file, index=False)

# ===========================
# Main
# ===========================
if __name__ == "__main__":
    with open(TICKERS_FILE, "r") as f:
        tickers = [line.strip().upper() for line in f if line.strip()]

    for ticker in tickers:
        try:
            print(f"\nProcessing {ticker} financials...")
            process_ticker(ticker)
            time.sleep(12)
        except Exception as e:
            print(f"Error processing {ticker}: {e}")

    for ticker in tickers:
        try:
            print(f"\nProcessing {ticker} dividends...")
            export_dividends(ticker)
        except Exception as e:
            print(f"Error processing {ticker}: {e}")


# In[ ]:




