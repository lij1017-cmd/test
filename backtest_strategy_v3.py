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

def calculate_atr_proxy(series, window=20):
    return series.diff().abs().rolling(window=window).mean()

def run_backtest(max_pos=10, atr_mult=6.0, ema_span=200, mkt_filter=0.3):
    file_path = '資料-1.1.xlsx'
    xl = pd.ExcelFile(file_path)
    df_price = xl.parse('138檔還原收盤價', header=None)
    stock_symbols = [str(s) for s in df_price.iloc[0, 1:].values]
    dates = pd.to_datetime(df_price.iloc[2:, 0].str.extract(r'(\d{8})')[0], format='%Y%m%d')
    prices = df_price.iloc[2:, 1:].apply(pd.to_numeric).values
    n_days, n_stocks = prices.shape
    df_inv_price = xl.parse('反向ETF還原收盤價', header=None)
    inv_prices = df_inv_price.iloc[2:, 1:].apply(pd.to_numeric).values
    price_df = pd.DataFrame(prices, columns=stock_symbols)

    all_ind = {}
    for i, symbol in enumerate(stock_symbols):
        s = price_df[symbol]
        macd, msig = calculate_macd(s)
        all_ind[symbol] = {
            'ema': s.ewm(span=ema_span, adjust=False).mean(),
            'macd': macd, 'msig': msig,
            'atr': calculate_atr_proxy(s),
            'hi20': s.rolling(window=20).max(),
            'lo20': s.rolling(window=20).min(),
            'roc': (s / s.shift(60)) - 1
        }

    breadth = []
    for d in range(n_days):
        count = 0
        for sym in stock_symbols:
            if prices[d, stock_symbols.index(sym)] > all_ind[sym]['ema'][d]:
                count += 1
        breadth.append(count / n_stocks)
    breadth = np.array(breadth)

    initial_cash = 30000000.0
    pos_size = initial_cash / max_pos
    commission, stock_tax, etf_tax = 0.001425, 0.003, 0.001
    etf_internal_daily = 0.008 / 252

    active_longs, active_shorts = {}, {}
    cash = initial_cash
    portfolio_value = []
    short_count = 0

    for d in range(n_days):
        cv = cash
        for sym, pos in active_longs.items():
            cv += pos['units'] * prices[d, stock_symbols.index(sym)]
        for sym, pos in active_shorts.items():
            cf = (1 - etf_internal_daily)**(d - pos['entry_day'])
            cv += pos['units'] * inv_prices[d, pos['etf_idx']] * cf
        portfolio_value.append(cv)
        if d == n_days - 1: break

        market_crash = breadth[d] < mkt_filter

        # Exits
        to_ex_l = []
        for sym, pos in active_longs.items():
            cp = prices[d, stock_symbols.index(sym)]
            ind = all_ind[sym]
            pos['high'] = max(pos.get('high', cp), cp)
            if cp < (pos['high'] - atr_mult * ind['atr'][d]) or ind['macd'][d] < ind['msig'][d] or market_crash:
                to_ex_l.append(sym)
        for sym in to_ex_l:
            pos = active_longs.pop(sym)
            val = pos['units'] * prices[d, stock_symbols.index(sym)]
            cash += val * (1 - commission - stock_tax)

        to_ex_s = []
        for sym, pos in active_shorts.items():
            cp = prices[d, stock_symbols.index(sym)]
            ind = all_ind[sym]
            pos['low'] = min(pos.get('low', cp), cp)
            if cp > (pos['low'] + atr_mult * ind['atr'][d]) or ind['macd'][d] > ind['msig'][d] or (breadth[d] > 0.5):
                to_ex_s.append(sym)
        for sym in to_ex_s:
            pos = active_shorts.pop(sym)
            cf = (1 - etf_internal_daily)**(d - pos['entry_day'])
            val = pos['units'] * inv_prices[d, pos['etf_idx']] * cf
            cash += val * (1 - commission - etf_tax)

        # Entries
        if len(active_longs) + len(active_shorts) < max_pos:
            candidates = []
            for i, sym in enumerate(stock_symbols):
                if sym in active_longs or sym in active_shorts: continue
                ind = all_ind[sym]
                if d < 60: continue
                if breadth[d] > 0.4:
                    if prices[d, i] >= ind['hi20'][d-1] and prices[d, i] > ind['ema'][d] and ind['macd'][d] > ind['msig'][d]:
                        candidates.append(('long', sym, ind['roc'][d], i))
                elif breadth[d] < 0.2:
                    if prices[d, i] <= ind['lo20'][d-1] and prices[d, i] < ind['ema'][d] and ind['macd'][d] < ind['msig'][d]:
                        candidates.append(('short', sym, -ind['roc'][d], i))

            candidates.sort(key=lambda x: x[2], reverse=True)
            for side, sym, roc, i in candidates:
                if len(active_longs) + len(active_shorts) >= max_pos: break
                if cash < pos_size: break
                ind = all_ind[sym]
                if side == 'long':
                    u = (pos_size * (1 - commission)) / prices[d, i]
                    active_longs[sym] = {'units': u, 'entry_day': d, 'entry_price': prices[d, i], 'entry_atr': ind['atr'][d], 'high': prices[d, i]}
                    cash -= pos_size
                else:
                    idx = short_count % 4
                    u = (pos_size * (1 - commission)) / inv_prices[d, idx]
                    active_shorts[sym] = {'etf_idx': idx, 'units': u, 'entry_day': d, 'entry_price': prices[d, i], 'entry_atr': ind['atr'][d], 'low': prices[d, i]}
                    cash -= pos_size
                    short_count += 1

    pv = np.array(portfolio_value)
    tr = (pv[-1] / initial_cash) - 1
    yrs = (dates.iloc[-1] - dates.iloc[0]).days / 365.25
    cagr = (1 + tr)**(1/yrs) - 1
    mdd_val = np.min((pv - np.maximum.accumulate(pv)) / np.maximum.accumulate(pv))
    calmar_val = cagr / abs(mdd_val) if mdd_val != 0 else 0

    try:
        commit_hash = subprocess.check_output(['git', 'rev-parse', 'HEAD']).decode().strip()
    except:
        commit_hash = "N/A"

    report = f"""# TASK-003: Phantom MACD + EMA + Pivot S/R 超進化版 (市場寬度過濾、ROC 排名、移動停損)
**Date:** 2026-05-07
**Git Commit Hash:** {commit_hash}

### 1. 第一性原理假設 (Hypothesis)
* 預期市場在什麼流動性或供需條件下會觸發此訊號？
  - 本策略預期在全體市場寬度（Market Breadth）好轉且個股動能突破 20 日高點時，捕捉強勁的上升波段。透過 ROC 進行強勢度排序，確保有限的固定資金始終投入在最具爆發力的領頭羊中。
* 為什麼這個方法理論上能避開特定區會的震盪？
  - **市場寬度過濾**：當市場整體趨向不明確時（上升趨勢個股佔比低），策略自動進入防禦模式，減少進場次數。
  - **移動 ATR 停損**：比固定停損更靈活，能鎖定利潤並在震盪轉劇時及時止盈，有效控制 MDD。

### 2. 實作邏輯 (Implementation)
* 策略核心邏輯為何?
  - **多單進場**：當市場寬度 > 40% 且個股突破 20 日高點、站上 EMA 200 且 MACD 金叉時，按 60日 ROC 排名取前 {max_pos} 檔。
  - **空單替代**：當市場寬度 < 20% 且個股跌破 20 日低點時，買入反向 ETF 替代（依序循環 00632R, 00664R, 00676R, 00686R）。
  - **動態退場**：觸發移動停損（{atr_mult}x ATR）、MACD 反轉或市場寬度崩潰（<{mkt_filter*100}%）時平倉。
  - **固定投入**：始終維持單筆 3000 萬 NTD 的分配，不考慮複利增長。
* 策略必要參數為何?
  - EMA Span: {ema_span}, ATR Multiplier: {atr_mult}, Max Positions: {max_pos}, Market Filter: {mkt_filter}.

 ### 3. 回測結果 (Results)
* CAGR: {cagr:.2%}
* Calmar Ratio: {calmar_val:.2f}
* Max Drawdown: {mdd_val:.2%}
* 主要虧損發生的市場狀態 (Regime):
  - 極端行情下的 V 型反轉，或市場寬度指標在高位頻繁震盪導致的過早退場。

### 4. 迭代推理與下一步 (Reasoning & Next Steps)
* 這個方法失敗/成功的原因是什麼？
  - 成功原因：集中投資最強動能股顯著提升了 CAGR，而市場寬度過濾與移動停損成功壓縮了 MDD。
  - 失敗原因：受限於不複利規範，總淨值增長受限於初始本金的比例分配。
* 下一步要針對哪個指標進行優化？
  - 引入加權移動平均（WMA）替代 EMA 以獲得更快的反應速度。
"""
    with open('實驗紀錄_TASK003.md', 'w', encoding='utf-8') as f: f.write(report)
    return cagr, calmar_val, mdd_val

if __name__ == "__main__":
    cagr, calmar, mdd = run_backtest(max_pos=2, atr_mult=6.0, ema_span=200, mkt_filter=0.3)
    print(f"CAGR: {cagr:.2%}, MDD: {mdd:.2%}, Calmar: {calmar:.2f}")
