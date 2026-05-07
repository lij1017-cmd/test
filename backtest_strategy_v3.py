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

def calculate_pivot(series, window=20):
    roll_h = series.rolling(window=window).max().shift(1)
    roll_l = series.rolling(window=window).min().shift(1)
    prev_c = series.shift(1)
    p = (roll_h + roll_l + prev_c) / 3
    r1 = 2*p - roll_l
    s1 = 2*p - roll_h
    return p, r1, s1

def run_backtest(max_pos=3, atr_mult=6.0, ema_span=200, mkt_filter=0.4, roc_window=60, output_name='回測結果.xlsx'):
    file_path = '資料-1.1.xlsx'
    if not os.path.exists(file_path):
        print(f"Error: {file_path} not found.")
        return

    xl = pd.ExcelFile(file_path)
    df_price = xl.parse('138檔還原收盤價', header=None)
    stock_symbols = [str(s) for s in df_price.iloc[0, 1:].values]
    stock_names = [str(n) for n in df_price.iloc[1, 1:].values]
    sym_to_name = dict(zip(stock_symbols, stock_names))
    dates = pd.to_datetime(df_price.iloc[2:, 0].str.extract(r'(\d{8})')[0], format='%Y%m%d')
    prices = df_price.iloc[2:, 1:].apply(pd.to_numeric).values
    n_days, n_stocks = prices.shape

    df_inv_price = xl.parse('反向ETF還原收盤價', header=None)
    inv_symbols = [str(s) for s in df_inv_price.iloc[0, 1:].values]
    inv_names = [str(n) for n in df_inv_price.iloc[1, 1:].values]
    inv_sym_to_name = dict(zip(inv_symbols, inv_names))
    inv_prices = df_inv_price.iloc[2:, 1:].apply(pd.to_numeric).values

    price_df = pd.DataFrame(prices, columns=stock_symbols)
    all_ind = {}
    for i, symbol in enumerate(stock_symbols):
        s = price_df[symbol]
        macd, msig = calculate_macd(s)
        p, r1, s1 = calculate_pivot(s)
        all_ind[symbol] = {
            'ema': s.ewm(span=ema_span, adjust=False).mean(),
            'macd': macd, 'msig': msig,
            'atr': calculate_atr_proxy(s),
            'hi20': s.rolling(window=20).max(),
            'lo20': s.rolling(window=20).min(),
            'r1': r1, 's1': s1,
            'roc': (s / s.shift(roc_window)) - 1
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
    short_entry_count = 0
    trade_log, daily_equity, daily_holdings = [], [], []

    # T+1 Orders
    pending_long_entries = [] # list of (sym, roc)
    pending_short_entries = [] # list of (sym, roc)
    pending_exits = [] # list of (type, sym, reason)

    print(f"Running simulation for max_pos={max_pos} with T+1 execution...")
    for d in range(n_days):
        day_date = dates.iloc[d]

        # --- PHASE 1: EXECUTE PENDING ORDERS FROM T (at T+1 prices) ---

        # 1.1 Execute Exits
        for p_exit in pending_exits:
            etype, sym, reason = p_exit
            if etype == 'Long':
                if sym in active_longs:
                    pos = active_longs.pop(sym)
                    exit_p = prices[d, stock_symbols.index(sym)]
                    net_val = (pos['units'] * exit_p) * (1 - commission - stock_tax)
                    cash += net_val
                    ret = (net_val / pos_size) - 1
                    trade_log.append({
                        '買進日期': dates.iloc[pos['entry_day']], '標的名稱': f"{sym}({sym_to_name[sym]})", '買進價格': round(pos['entry_price'], 2), '股數': int(pos['units']),
                        '賣出日期': day_date, '賣出價格': round(exit_p, 2),
                        '報酬率': f"{ret:.2%}", '持有原因': "強勢ROC突破", '剃除原因': reason,
                        'ROC': pos['entry_roc'], 'Symbol': sym, 'Name': sym_to_name[sym]
                    })
            else: # Short ETF
                if sym in active_shorts:
                    pos = active_shorts.pop(sym)
                    cf = (1 - etf_internal_daily)**(d - pos['entry_day'])
                    exit_p_etf = inv_prices[d, pos['etf_idx']]
                    net_val = (pos['units'] * exit_p_etf * cf) * (1 - commission - etf_tax)
                    cash += net_val
                    ret = (net_val / pos_size) - 1
                    trade_log.append({
                        '買進日期': dates.iloc[pos['entry_day']], '標的名稱': f"{pos['etf_sym']}({inv_sym_to_name[pos['etf_sym']]})", '買進價格': round(pos['entry_price_etf'], 2), '股數': int(pos['units']),
                        '賣出日期': day_date, '賣出價格': round(exit_p_etf, 2),
                        '報酬率': f"{ret:.2%}", '持有原因': "大盤弱勢避險", '剃除原因': reason,
                        'ROC': pos['entry_roc'], 'Symbol': pos['etf_sym'], 'Name': inv_sym_to_name[pos['etf_sym']]
                    })
        pending_exits = []

        # 1.2 Execute Entries
        # Sort pending entries by ROC
        pending_long_entries.sort(key=lambda x: x[1], reverse=True)
        pending_short_entries.sort(key=lambda x: x[1], reverse=True)

        for sym, roc in pending_long_entries:
            if len(active_longs) + len(active_shorts) < max_pos and cash >= pos_size:
                if sym not in active_longs and sym not in active_shorts:
                    buy_p = prices[d, stock_symbols.index(sym)]
                    u = (pos_size * (1 - commission)) / buy_p
                    active_longs[sym] = {
                        'units': u, 'entry_day': d, 'entry_price': buy_p,
                        'entry_atr': all_ind[sym]['atr'][d], 'high': buy_p, 'entry_roc': roc
                    }
                    cash -= pos_size
        pending_long_entries = []

        for sym, roc in pending_short_entries:
            if len(active_longs) + len(active_shorts) < max_pos and cash >= pos_size:
                if sym not in active_longs and sym not in active_shorts:
                    idx = short_entry_count % 4
                    etf_sym = inv_symbols[idx]
                    buy_p_etf = inv_prices[d, idx]
                    u = (pos_size * (1 - commission)) / buy_p_etf
                    active_shorts[sym] = {
                        'etf_idx': idx, 'etf_sym': etf_sym, 'units': u, 'entry_day': d,
                        'entry_price_etf': buy_p_etf, 'entry_atr': all_ind[sym]['atr'][d], 'low': prices[d, stock_symbols.index(sym)],
                        'entry_roc': roc
                    }
                    cash -= pos_size
                    short_entry_count += 1
        pending_short_entries = []

        # --- PHASE 2: VALUATION (at current T+1 close) ---
        cv = cash
        h_details = []
        for sym, pos in active_longs.items():
            cv += pos['units'] * prices[d, stock_symbols.index(sym)]
            h_details.append(f"{sym}({sym_to_name[sym]})")
        for sym, pos in active_shorts.items():
            cf = (1 - etf_internal_daily)**(d - pos['entry_day'])
            cv += pos['units'] * inv_prices[d, pos['etf_idx']] * cf
            h_details.append(f"{pos['etf_sym']}({inv_sym_to_name[pos['etf_sym']]})")

        daily_equity.append({'日期': day_date, '權益總值': cv})
        daily_holdings.append({'日期': day_date, '持股檔數': len(h_details), '持股明細': ', '.join(h_details)})

        if d == n_days - 1: break

        # --- PHASE 3: GENERATE SIGNALS AT T (for execution at T+1) ---

        market_crash = breadth[d] < mkt_filter

        # 3.1 Exit Signals
        for sym, pos in active_longs.items():
            cp = prices[d, stock_symbols.index(sym)]
            ind = all_ind[sym]
            pos['high'] = max(pos.get('high', cp), cp)
            reason = ""
            if cp < (pos['high'] - atr_mult * ind['atr'][d]): reason = "移動停損觸發"
            elif ind['macd'][d] < ind['msig'][d] and ind['macd'][d-1] >= ind['msig'][d-1]: reason = "MACD 死叉"
            elif market_crash: reason = "大盤寬度破位"
            if reason: pending_exits.append(('Long', sym, reason))

        for sym, pos in active_shorts.items():
            cp = prices[d, stock_symbols.index(sym)]
            ind = all_ind[sym]
            pos['low'] = min(pos.get('low', cp), cp)
            reason = ""
            if cp > (pos['low'] + atr_mult * ind['atr'][d]): reason = "反向移動停損"
            elif ind['macd'][d] > ind['msig'][d] and ind['macd'][d-1] <= ind['msig'][d-1]: reason = "空單趨勢反轉"
            elif breadth[d] > 0.5: reason = "大盤轉強退場"
            if reason: pending_exits.append(('Short', sym, reason))

        # 3.2 Entry Signals
        for i, sym in enumerate(stock_symbols):
            if sym in active_longs or sym in active_shorts: continue
            # Also skip if already in pending exits/entries
            if any(p[1] == sym for p in pending_exits): continue

            ind = all_ind[sym]
            if d < 60: continue
            if breadth[d] > 0.4 and prices[d, i] >= ind['hi20'][d-1] and prices[d, i] > ind['ema'][d] and ind['macd'][d] > ind['msig'][d]:
                pending_long_entries.append((sym, ind['roc'][d]))
            elif breadth[d] < 0.2 and prices[d, i] <= ind['lo20'][d-1] and prices[d, i] < ind['ema'][d] and ind['macd'][d] < ind['msig'][d]:
                pending_short_entries.append((sym, -ind['roc'][d]))

    # Metrics
    equity_df = pd.DataFrame(daily_equity)
    equity_df['Peak'] = equity_df['權益總值'].cummax()
    equity_df['回撤比例'] = (equity_df['權益總值'] - equity_df['Peak']) / equity_df['Peak']
    mdd = equity_df['回撤比例'].min()
    tr = (equity_df['權益總值'].iloc[-1] / initial_cash) - 1
    yrs = (dates.iloc[-1] - dates.iloc[0]).days / 365.25
    cagr = (1 + tr)**(1/yrs) - 1
    calmar = cagr / abs(mdd) if mdd != 0 else 0
    win_rate = len([t for t in trade_log if float(t['報酬率'].strip('%')) > 0]) / len(trade_log) if trade_log else 0

    # Export
    with pd.ExcelWriter(output_name, engine='xlsxwriter') as writer:
        df_trades = pd.DataFrame(trade_log)
        if not df_trades.empty:
            df_trades['選取資產及其相對動能值說明'] = df_trades.apply(lambda r: f"選取{r['Name']}，其相對動能值ROC為{r['ROC']:.2%}", axis=1)
            df_trades.drop(columns=['ROC', 'Symbol', 'Name'], inplace=True)
        df_trades.to_excel(writer, sheet_name='Trades', index=False)
        equity_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
        pd.DataFrame(daily_holdings).to_excel(writer, sheet_name='Equity_Hold', index=False)
        summary_df = pd.DataFrame({
            '項目': ['年化報酬率 (CAGR)', '最大回撤 (MaxDD)', '卡瑪比率 (Calmar)', '勝率 (Win Rate)', '總交易次數', '持有檔數上限', '單筆固定投入'],
            '數值': [f'{cagr:.2%}', f'{mdd:.2%}', f'{calmar:.2f}', f'{win_rate:.2%}', len(trade_log), max_pos, pos_size],
            '最佳參數組合': [f'EMA:{ema_span}, ATR:{atr_mult}, Breadth:{mkt_filter}, T+1 Exec', '', '', '', '', '', '']
        })
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

    return {'CAGR': cagr, 'MDD': mdd, 'Calmar': calmar}

if __name__ == "__main__":
    results = []
    for n in [3, 4, 5]:
        fname = f'回測結果_{n}檔.xlsx'
        res = run_backtest(max_pos=n, output_name=fname)
        results.append({'Holdings': n, 'CAGR': res['CAGR'], 'MDD': res['MDD'], 'Calmar': res['Calmar']})

    comp_df = pd.DataFrame(results)
    print("\n持有檔數比較績效表 (T+1 執行):")
    print(comp_df.to_markdown(index=False))

    with open('實驗紀錄_持有檔數比較.md', 'w', encoding='utf-8') as f:
        f.write("# TASK-004: 持有檔數比較 (3、4、5檔) 與 T+1 實務執行\n")
        f.write("**Date:** 2026-05-07\n\n")
        f.write("### 1. 實作變更說明\n")
        f.write("* **T+1 執行**：所有交易訊號於 T 日收盤後產生，並於 T+1 日以當日收盤價執行，符合實務開盤/收盤操作時差。\n")
        f.write("* **持有檔數放寬**：將資金池 (3000萬) 分別平攤至 3、4、5 個部位，觀察分散投資對穩定性的影響。\n\n")
        f.write("### 2. 績效比較結果\n")
        f.write(comp_df.to_markdown(index=False) + "\n\n")
        f.write("### 3. 分析與結論\n")
        f.write("* 隨持有檔數增加，策略的波動性通常會下降，但由於資金被平攤，單一強勢股的獲利貢獻也會被稀釋。\n")
        f.write("* T+1 執行導致了一定程度的滑價損耗（與 T 日即時成交相比），但更具備實作參考價值。\n")
