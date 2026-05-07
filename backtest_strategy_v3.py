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

def run_backtest(max_pos=2, atr_mult=6.0, ema_span=200, mkt_filter=0.4, roc_window=60):
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

    for d in range(n_days):
        cv = cash
        h_details = []
        for sym, pos in active_longs.items():
            cv += pos['units'] * prices[d, stock_symbols.index(sym)]
            h_details.append(f"{sym}({sym_to_name[sym]})")
        for sym, pos in active_shorts.items():
            cf = (1 - etf_internal_daily)**(d - pos['entry_day'])
            cv += pos['units'] * inv_prices[d, pos['etf_idx']] * cf
            h_details.append(f"{pos['etf_sym']}({inv_sym_to_name[pos['etf_sym']]})")

        daily_equity.append({'日期': dates.iloc[d], '權益總值': cv})
        daily_holdings.append({'日期': dates.iloc[d], '持股檔數': len(h_details), '持股明細': ', '.join(h_details)})

        if d == n_days - 1: break

        # Exits
        to_ex_l = []
        for sym, pos in active_longs.items():
            cp = prices[d, stock_symbols.index(sym)]
            ind = all_ind[sym]
            pos['high'] = max(pos.get('high', cp), cp)
            reason = ""
            if cp < (pos['high'] - atr_mult * ind['atr'][d]): reason = "移動停損觸發"
            elif ind['macd'][d] < ind['msig'][d] and ind['macd'][d-1] >= ind['msig'][d-1]: reason = "MACD 死叉"
            elif breadth[d] < mkt_filter: reason = "大盤寬度破位"
            if reason: to_ex_l.append((sym, reason))

        for sym, reason in to_ex_l:
            pos = active_longs.pop(sym)
            exit_p = prices[d, stock_symbols.index(sym)]
            net_val = (pos['units'] * exit_p) * (1 - commission - stock_tax)
            cash += net_val
            ret = (net_val / pos_size) - 1
            trade_log.append({
                '買進日期': dates.iloc[pos['entry_day']], '標的名稱': f"{sym}({sym_to_name[sym]})", '買進價格': round(pos['entry_price'], 2), '買進股數': int(pos['units']),
                '賣出日期': dates.iloc[d], '賣出剃除商品': f"{sym}({sym_to_name[sym]})", '賣出價格': round(exit_p, 2), '賣出股數': int(pos['units']),
                '每期報酬表現': f"{ret:.2%}", '持有原因': "強勢ROC突破", '剃除原因': reason,
                '最佳參數': f'EMA{ema_span}/ATR{atr_mult}', 'ROC': pos['entry_roc'], 'Symbol': sym
            })

        to_ex_s = []
        for sym, pos in active_shorts.items():
            cp = prices[d, stock_symbols.index(sym)]
            ind = all_ind[sym]
            pos['low'] = min(pos.get('low', cp), cp)
            reason = ""
            if cp > (pos['low'] + atr_mult * ind['atr'][d]): reason = "反向移動停損"
            elif ind['macd'][d] > ind['msig'][d] and ind['macd'][d-1] <= ind['msig'][d-1]: reason = "空單趨勢反轉"
            elif breadth[d] > 0.5: reason = "大盤轉強退場"
            if reason: to_ex_s.append((sym, reason))

        for sym, reason in to_ex_s:
            pos = active_shorts.pop(sym)
            cf = (1 - etf_internal_daily)**(d - pos['entry_day'])
            exit_p_etf = inv_prices[d, pos['etf_idx']]
            net_val = (pos['units'] * exit_p_etf * cf) * (1 - commission - etf_tax)
            cash += net_val
            ret = (net_val / pos_size) - 1
            trade_log.append({
                '買進日期': dates.iloc[pos['entry_day']], '標的名稱': f"{pos['etf_sym']}({inv_sym_to_name[pos['etf_sym']]})", '買進價格': round(pos['entry_price_etf'], 2), '買進股數': int(pos['units']),
                '賣出日期': dates.iloc[d], '賣出剃除商品': f"{pos['etf_sym']}({inv_sym_to_name[pos['etf_sym']]})", '賣出價格': round(exit_p_etf, 2), '賣出股數': int(pos['units']),
                '每期報酬表現': f"{ret:.2%}", '持有原因': "弱勢大盤避險", '剃除原因': reason,
                '最佳參數': f'EMA{ema_span}/ATR{atr_mult}', 'ROC': pos['entry_roc'], 'Symbol': pos['etf_sym']
            })

        # Entries
        if len(active_longs) + len(active_shorts) < max_pos:
            candidates = []
            for i, sym in enumerate(stock_symbols):
                if sym in active_longs or sym in active_shorts: continue
                ind = all_ind[sym]
                if d < 60: continue
                if breadth[d] > 0.4 and prices[d, i] >= ind['hi20'][d-1] and prices[d, i] > ind['ema'][d] and ind['macd'][d] > ind['msig'][d]:
                    candidates.append(('long', sym, ind['roc'][d], i))
                elif breadth[d] < 0.2 and prices[d, i] <= ind['lo20'][d-1] and prices[d, i] < ind['ema'][d] and ind['macd'][d] < ind['msig'][d]:
                    candidates.append(('short', sym, -ind['roc'][d], i))

            candidates.sort(key=lambda x: x[2], reverse=True)
            for side, sym, roc, i in candidates:
                if len(active_longs) + len(active_shorts) >= max_pos: break
                if cash < pos_size: break
                ind = all_ind[sym]
                if side == 'long':
                    u = (pos_size * (1 - commission)) / prices[d, i]
                    active_longs[sym] = {'units': u, 'entry_day': d, 'entry_price': prices[d, i], 'entry_atr': ind['atr'][d], 'high': prices[d, i], 'entry_roc': roc}
                    cash -= pos_size
                else:
                    idx = short_entry_count % 4
                    etf_sym = inv_symbols[idx]
                    u = (pos_size * (1 - commission)) / inv_prices[d, idx]
                    active_shorts[sym] = {'etf_idx': idx, 'etf_sym': etf_sym, 'units': u, 'entry_day': d, 'entry_price_etf': inv_prices[d, idx], 'entry_atr': ind['atr'][d], 'low': prices[d, i], 'entry_roc': roc}
                    cash -= pos_size
                    short_entry_count += 1

    equity_df = pd.DataFrame(daily_equity)
    equity_df['Peak'] = equity_df['權益總值'].cummax()
    equity_df['回撤比例'] = (equity_df['權益總值'] - equity_df['Peak']) / equity_df['Peak']
    mdd = equity_df['回撤比例'].min()
    tr = (equity_df['權益總值'].iloc[-1] / initial_cash) - 1
    yrs = (dates.iloc[-1] - dates.iloc[0]).days / 365.25
    cagr = (1 + tr)**(1/yrs) - 1
    calmar = cagr / abs(mdd) if mdd != 0 else 0
    win_rate = len([t for t in trade_log if float(t['每期報酬表現'].strip('%')) > 0]) / len(trade_log) if trade_log else 0

    output_file = '回測結果彙整.xlsx'
    print(f"Generating Excel: {output_file}")
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # 1. Trades
        df_trades = pd.DataFrame(trade_log)
        if not df_trades.empty:
            df_trades['選取資產及其相對動能值說明'] = df_trades.apply(lambda r: f"選取{r['標的名稱']}，其相對動能值ROC為{r['ROC']:.2%}", axis=1)
            df_trades.drop(columns=['ROC', 'Symbol'], inplace=True)
        df_trades.to_excel(writer, sheet_name='Trades', index=False)

        # 2. Equity_Curve
        equity_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
        worksheet = writer.sheets['Equity_Curve']
        chart = writer.book.add_chart({'type': 'line'})
        chart.add_series({
            'name': 'Equity Curve',
            'categories': ['Equity_Curve', 1, 0, len(equity_df), 0],
            'values': ['Equity_Curve', 1, 1, len(equity_df), 1]
        })
        chart.set_title({'name': '每日權益變化 (Equity Curve)'})
        worksheet.insert_chart('G2', chart)

        # 3. Equity_Hold
        pd.DataFrame(daily_holdings).to_excel(writer, sheet_name='Equity_Hold', index=False)

        # 4. Summary
        summary_df = pd.DataFrame({
            '項目': ['年化報酬率 (CAGR)', '最大回撤 (MaxDD)', '卡瑪比率 (Calmar)', '勝率 (Win Rate)', '總交易次數', '最佳參數組合'],
            '數值': [f'{cagr:.2%}', f'{mdd:.2%}', f'{calmar:.2f}', f'{win_rate:.2%}', len(trade_log), f'EMA:{ema_span}, ATR:{atr_mult}, MaxPos:{max_pos}, Breadth:{mkt_filter}']
        })
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    print("Success.")

if __name__ == "__main__":
    run_backtest(max_pos=2, atr_mult=6.0, ema_span=200, mkt_filter=0.4)
