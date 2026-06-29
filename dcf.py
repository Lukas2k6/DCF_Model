import yfinance as yf
import pandas as pd

def get_financials(ticker):
    stock = yf.Ticker(ticker)
    
    income_stmt = stock.financials
    balance_sheet = stock.balance_sheet
    cash_flow = stock.cashflow
    info = stock.info
    
    return {
        "income_statement": income_stmt,
        "balance_sheet": balance_sheet,
        "cash_flow": cash_flow,
        "info": info
    }

if __name__ == "__main__":
    ticker = input("Enter ticker: ").upper()
    data = get_financials(ticker)
    print(data["income_statement"])