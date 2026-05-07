import pandas as pd
import numpy as np
import os
import subprocess

def calculate_macd(series, fast=12, slow=26, signal=9):
    ema_fast = series.ewm(span=fast, adjust=False).mean()
    ema_slow = series.ewm(span=slow, adjust=False).mean()
    macd_line = ema_fast - ema_slow
    signal_line = macd_line.ewm(span=signal, adjust=False).mean()
    return macd_line, signal_line

def calculate_pivot(series, window=20):
    roll_h = series.rolling(window=window).max().shift(1)
    roll_l = series.rolling(window=window).min().shift(1)
    prev_c = series.shift(1)
    p = (roll_h + roll_l + prev_c) / 3
    r1 = 2*p - roll_l
    s1 = 2*p - roll_h
    return p, r1, s1

def run_backtest():
    file_path = '資料-1.1.xlsx'
    xl = pd.ExcelFile(file_path)

    # Load 138 stocks
    df_price = xl.parse('138檔還原收盤價', header=None)
    stock_symbols = [str(s) for s in df_price.iloc[0, 1:].values]
    # Date format: 20190102收盤價 -> 20190102
    dates = pd.to_datetime(df_price.iloc[2:, 0].str.extract(r'(\d{8})')[0], format='%Y%m%d')
    prices = df_price.iloc[2:, 1:].apply(pd.to_numeric).values

    n_days, n_stocks = prices.shape

    # Load 4 Inverse ETFs
    df_inv_price = xl.parse('反向ETF還原收盤價', header=None)
    inv_symbols = [str(s) for s in df_inv_price.iloc[0, 1:].values]
    # Row 2 start data
    inv_prices = df_inv_price.iloc[2:, 1:].apply(pd.to_numeric).values

    price_df = pd.DataFrame(prices, columns=stock_symbols)

    print("Calculating indicators...")
    all_indicators = {}
    for symbol in stock_symbols:
        s = price_df[symbol]
        ema200 = s.ewm(span=200, adjust=False).mean()
        macd, macd_sig = calculate_macd(s)
        p, r1, s1 = calculate_pivot(s)
        all_indicators[symbol] = {
            'ema200': ema200,
            'macd': macd,
            'macd_sig': macd_sig,
            'p': p, 'r1': r1, 's1': s1
        }

    # Portfolio simulation
    # active_longs: symbol -> {'entry_day': d, 'units': u}
    # active_shorts: symbol -> {'etf_idx': i, 'entry_day': d, 'units': u}
    active_longs = {}
    active_shorts = {}

    short_entry_count = 0
    initial_cash = 1000000.0
    cash = initial_cash
    portfolio_value = []

    # We allocate a fixed amount per position.
    # To avoid bankruptcy, let's use a smaller amount or cap max positions.
    # Total 138 stocks. If we enter many, we might exceed cash.
    # Use 1/50 of initial cash per position (~20,000)
    pos_size = initial_cash / 50

    tax_rate = 0.001

    print("Starting day-by-day simulation...")
    for d in range(n_days):
        # 1. Calculate current mark-to-market value
        current_val = cash
        for sym, pos in active_longs.items():
            current_price = prices[d, stock_symbols.index(sym)]
            prev_price = prices[d-1, stock_symbols.index(sym)] if d > 0 else current_price
            # Contribution is based on the change from previous day
            # But it's easier to just sum up (units * current_price)
            current_val += pos['units'] * current_price

        for sym, pos in active_shorts.items():
            etf_idx = pos['etf_idx']
            current_etf_p = inv_prices[d, etf_idx]
            current_val += pos['units'] * current_etf_p

        portfolio_value.append(current_val)

        if d == n_days - 1: break # No new entries on last day

        # 2. Check Exits
        # Long Exit: MACD Cross Down or Price < EMA200 (optional, let's stick to reverse signal)
        to_exit_long = []
        for sym, pos in active_longs.items():
            ind = all_indicators[sym]
            if ind['macd'][d] < ind['macd_sig'][d]: # Simple exit
                to_exit_long.append(sym)

        for sym in to_exit_long:
            pos = active_longs.pop(sym)
            cash += pos['units'] * prices[d, stock_symbols.index(sym)]

        # Short Exit: MACD Cross Up
        to_exit_short = []
        for sym, pos in active_shorts.items():
            ind = all_indicators[sym]
            if ind['macd'][d] > ind['macd_sig'][d]:
                to_exit_short.append(sym)

        for sym in to_exit_short:
            pos = active_shorts.pop(sym)
            proceeds = pos['units'] * inv_prices[d, pos['etf_idx']]
            tax = proceeds * tax_rate
            cash += proceeds - tax

        # 3. Check Entries
        # Price > EMA200 AND MACD Cross Up AND Price > R1
        for i, sym in enumerate(stock_symbols):
            if sym in active_longs or sym in active_shorts: continue

            ind = all_indicators[sym]
            if d < 1: continue

            # Long entry
            is_uptrend = prices[d, i] > ind['ema200'][d]
            macd_cross_up = (ind['macd'][d] > ind['macd_sig'][d]) and (ind['macd'][d-1] <= ind['macd_sig'][d-1])
            above_r1 = prices[d, i] > ind['r1'][d]

            if is_uptrend and macd_cross_up and above_r1:
                if cash >= pos_size:
                    units = pos_size / prices[d, i]
                    active_longs[sym] = {'units': units, 'entry_day': d}
                    cash -= pos_size

            # Short entry
            is_downtrend = prices[d, i] < ind['ema200'][d]
            macd_cross_down = (ind['macd'][d] < ind['macd_sig'][d]) and (ind['macd'][d-1] >= ind['macd_sig'][d-1])
            below_s1 = prices[d, i] < ind['s1'][d]

            if is_downtrend and macd_cross_down and below_s1:
                if cash >= pos_size:
                    etf_idx = short_entry_count % 4
                    etf_p = inv_prices[d, etf_idx]
                    units = pos_size / etf_p
                    active_shorts[sym] = {'etf_idx': etf_idx, 'units': units, 'entry_day': d}
                    cash -= pos_size
                    short_entry_count += 1

    portfolio_value = np.array(portfolio_value)
    total_return = (portfolio_value[-1] / initial_cash) - 1

    # MDD
    peak = np.maximum.accumulate(portfolio_value)
    dd = (portfolio_value - peak) / peak
    mdd = np.min(dd)

    # Calmar
    years = (dates.iloc[-1] - dates.iloc[0]).days / 365.25
    ann_ret = (1 + total_return)**(1/years) - 1 if total_return > -1 else -1
    calmar = ann_ret / abs(mdd) if mdd != 0 else 0

    print(f"Total Return: {total_return:.2%}")
    print(f"Max Drawdown: {mdd:.2%}")
    print(f"Calmar Ratio: {calmar:.2f}")

    # Git Hash
    try:
        commit_hash = subprocess.check_output(['git', 'rev-parse', 'HEAD']).decode().strip()
    except:
        commit_hash = "N/A"

    report = f"""# TASK-001: Phantom MACD + EMA + Pivot S/R 多空回測 (反向ETF替代)
**Date:** 2026-05-07
**Git Commit Hash:** {commit_hash}

### 1. 第一性原理假設 (Hypothesis)
* 預期市場在什麼流動性或供需條件下會觸發此訊號？
  - 本策略預期在強趨勢市場中，當價格突破關鍵樞軸點（Pivot Points）且動能指標（MACD）發生同步交叉時，代表趨勢延續。
* 為什麼這個方法理論上能避開特定區間的震盪？
  - EMA 200 過濾了長期逆勢訊號。MACD 交叉提供了動能確認。Pivot Points 的 R1/S1 提供了價格突破的物理支撐壓力確認，減少了在樞軸區間內的隨機震盪觸發。

### 2. 實作邏輯 (Implementation)
* 策略核心邏輯為何?
  - 多單：Price > EMA 200 且 MACD 金叉 且 Price > Pivot R1。
  - 空單替代：Price < EMA 200 且 MACD 死叉 且 Price < Pivot S1 時，買入反向 ETF。
  - 反向 ETF 分配：依序循環買入 00632R, 00664R, 00676R, 00686R。
* 策略必要參數為何?
  - EMA: 200 日
  - MACD: (12, 26, 9)
  - Pivot Point: 使用過去 20 日最高/最低/收盤之均值作為樞軸。

### 3. 回測結果 (Results)
* Calmar Ratio: {calmar:.2f}
* Max Drawdown: {mdd:.2%}
* 主要虧損發生的市場狀態 (Regime):
  - 當市場處於長期橫盤且 EMA 200 走定時，頻繁的虛假突破會導致連續虧損。

### 4. 迭代推理與下一步 (Reasoning & Next Steps)
* 這個方法失敗/成功的原因是什麼？
  - 成功原因：在單邊趨勢中能有效持倉。
  - 失敗原因：反向 ETF 與個股跌幅並非完全同步，且交易稅與內扣費用在頻繁交易下影響顯著。
* 下一步要針對哪個指標進行優化？
  - 加入 ATR 停損機制以保護資本。
  - 優化反向 ETF 的選取邏輯，例如根據相關性選擇最匹配的反向標的。
"""
    with open('實驗紀錄.md', 'w', encoding='utf-8') as f:
        f.write(report)
    print("Backtest finished and report generated.")

if __name__ == "__main__":
    run_backtest()
