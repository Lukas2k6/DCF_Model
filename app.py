from flask import Flask, jsonify, render_template
import yfinance as yf
import pandas as pd
import os

app = Flask(__name__, template_folder=r'C:\Users\lukas')

from flask_cors import CORS
CORS(app)

def safe_get(df, keys):
    for key in keys:
        try:
            val = df.loc[key].iloc[0]
            if val is not None and not pd.isna(val):
                return float(val)
        except:
            continue
    return 0.0

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/dcf/<ticker>')
def get_dcf(ticker):
    try:
        stock = yf.Ticker(ticker.upper())
        info = stock.info

        if not info or (info.get("regularMarketPrice") is None and info.get("currentPrice") is None):
            return jsonify({"error": f"{ticker.upper()} is not an existing ticker, sorry!"}), 400
        cf    = stock.cashflow
        if cf is None or cf.empty:
            return jsonify({"error": f"No cash flow data available for {ticker.upper()}"}), 400
        bs    = stock.balance_sheet
        if bs is None or bs.empty:
            return jsonify({"error": f"No balance sheet data available for {ticker.upper()}"}), 400
        inc   = stock.financials

        op_cf      = safe_get(cf, ["Operating Cash Flow", "Total Cash From Operating Activities"])
        capex      = safe_get(cf, ["Capital Expenditure", "Capital Expenditures"])
        fcf        = op_cf + capex
        if fcf <= 0:
            return jsonify({"error": f"{ticker.upper()} has no positive FCF — DCF model not applicable. Try something other than a shell company please."}), 400
        total_debt = safe_get(bs, ["Total Debt", "Long Term Debt", "Short Long Term Debt"])
        cash       = safe_get(bs, ["Cash And Cash Equivalents", "Cash", "Cash And Short Term Investments"])
        shares     = float(info.get("sharesOutstanding") or info.get("impliedSharesOutstanding") or 0) / 1e6
        if shares <= 0:
            return jsonify({"error": f"Could not retrieve share count for {ticker.upper()}"}), 400
        cur_price  = float(info.get("currentPrice") or info.get("regularMarketPrice") or 0)
        if cur_price <= 0:
            return jsonify({"error": f"Could not retrieve current price for {ticker.upper()}"}), 400
        beta       = float(info.get("beta") or 1.0)

        revs = inc.loc["Total Revenue"].dropna() if "Total Revenue" in inc.index else None
        rev_growth = float((revs.iloc[0] / revs.iloc[1]) - 1) if revs is not None and len(revs) >= 2 else 0.10

        return jsonify({
            "ticker":     ticker.upper(),
            "price":      cur_price,
            "shares":     shares,
            "beta":       beta,
            "fcf":        fcf / 1e6,
            "op_cf":      op_cf / 1e6,
            "capex":      capex / 1e6,
            "total_debt": total_debt / 1e6,
            "cash":       cash / 1e6,
            "rev_growth": round(rev_growth, 4)
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 400

if __name__ == '__main__':
    app.run(debug=True)
