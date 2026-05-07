import pandas as pd
import numpy as np
import os

def load_and_clean_data(file_path):
    print(f"Loading data from {file_path}...")
    xl = pd.ExcelFile(file_path)
    df = xl.parse('138檔還原收盤價', header=None)

    # Symbols in row 0, starting from col 1
    symbols = df.iloc[0, 1:].tolist()
    # Names in row 1, starting from col 1
    names = df.iloc[1, 1:].tolist()

    # Dates in col 0, starting from row 2. Extract 8 digits.
    date_col = df.iloc[2:, 0].astype(str)
    dates = pd.to_datetime(date_col.str.extract(r'(\d{8})')[0], format='%Y%m%d')

    # Prices starting from row 2, col 1
    prices = df.iloc[2:, 1:].apply(pd.to_numeric).values

    price_df = pd.DataFrame(prices, index=dates, columns=symbols)

    # Cleaning data:
    # 1. Back-fill leading gaps from the first valid data point
    # 2. Forward-fill internal gaps with the last available value
    # Pandas bfill and ffill can do this.
    # Specifically: ffill() followed by bfill() or vice versa.
    # The requirement: "back-filling leading gaps from the first valid data point and forward-fill internal gaps"
    # Usually this means ffill first (to fill internal gaps from previous values),
    # then bfill (to fill the very beginning if it starts with NaN).

    cleaned_df = price_df.ffill().bfill()

    return cleaned_df

def calculate_indicators(df):
    print("Calculating SMA(200) and RSI(14)...")
    indicators = {}
    for symbol in df.columns:
        series = df[symbol]
        sma200 = series.rolling(window=200).mean()

        # RSI calculation
        delta = series.diff()
        gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
        # Wilder's smoothing is often used for RSI, but standard rolling mean is a common approximation.
        # QuantConnect often uses Wilder's. Let's use Wilder's (ewm) if possible.

        avg_gain = delta.where(delta > 0, 0).ewm(alpha=1/14, adjust=False).mean()
        avg_loss = (-delta.where(delta < 0, 0)).ewm(alpha=1/14, adjust=False).mean()

        rs = avg_gain / avg_loss
        rsi = 100 - (100 / (1 + rs))

        indicators[symbol] = pd.DataFrame({
            'sma200': sma200,
            'rsi': rsi
        }, index=df.index)
    return indicators

def run_backtest(price_df, indicators_dict, initial_cash=30000000.0, num_slots=10):
    print(f"Running backtest with {initial_cash} TWD and {num_slots} slots...")
    cash = initial_cash
    pos_size = initial_cash / num_slots

    # Active positions: symbol -> {units, entry_price, entry_date}
    active_positions = {}

    portfolio_history = []
    dates = price_df.index

    buy_fee_rate = 0.001425
    sell_fee_rate = 0.001425
    tax_rate = 0.003

    for i, date in enumerate(dates):
        # 1. Update Portfolio Value (Mark-to-Market)
        current_market_value = 0
        for symbol, pos in active_positions.items():
            current_price = price_df.loc[date, symbol]
            current_market_value += pos['units'] * current_price

        portfolio_history.append(cash + current_market_value)

        if i == len(dates) - 1:
            break

        # 2. Check for Exits
        to_exit = []
        for symbol, pos in active_positions.items():
            price = price_df.loc[date, symbol]
            sma = indicators_dict[symbol].loc[date, 'sma200']
            rsi = indicators_dict[symbol].loc[date, 'rsi']

            # Exit conditions: RSI < 70 or Price < SMA200
            if rsi < 70 or price < sma:
                to_exit.append(symbol)

        for symbol in to_exit:
            pos = active_positions.pop(symbol)
            exit_price = price_df.loc[date, symbol]
            gross_proceeds = pos['units'] * exit_price
            fees = gross_proceeds * sell_fee_rate
            tax = gross_proceeds * tax_rate
            cash += (gross_proceeds - fees - tax)
            # print(f"[{date.date()}] EXIT {symbol} at {exit_price:.2f}")

        # 3. Check for Entries
        if len(active_positions) < num_slots:
            available_slots = num_slots - len(active_positions)

            # Potential entries
            candidates = []
            for symbol in price_df.columns:
                if symbol in active_positions:
                    continue

                price = price_df.loc[date, symbol]
                sma = indicators_dict[symbol].loc[date, 'sma200']
                rsi = indicators_dict[symbol].loc[date, 'rsi']

                # Check for NaNs
                if pd.isna(sma) or pd.isna(rsi):
                    continue

                # Entry conditions: Price > SMA200 AND RSI > 70
                # We also want to see it JUST crossing or staying above.
                # Strategy: "RSI Overbought Continuation". Usually means RSI > 70.
                if price > sma and rsi > 70:
                    candidates.append(symbol)

            # If many candidates, maybe sort by RSI or just take the first available
            for symbol in candidates:
                if len(active_positions) >= num_slots:
                    break

                entry_price = price_df.loc[date, symbol]
                # Fixed investment amount (pos_size)
                # But we need to account for buy fees
                # Amount spent (including fee) = pos_size
                # pos_size = units * entry_price * (1 + buy_fee_rate)
                units = pos_size / (entry_price * (1 + buy_fee_rate))

                if cash >= (units * entry_price * (1 + buy_fee_rate)):
                    active_positions[symbol] = {
                        'units': units,
                        'entry_price': entry_price,
                        'entry_date': date
                    }
                    cash -= (units * entry_price * (1 + buy_fee_rate))
                    # print(f"[{date.date()}] ENTRY {symbol} at {entry_price:.2f}")

    portfolio_history = pd.Series(portfolio_history, index=dates)
    return portfolio_history

def analyze_performance(history):
    total_return = (history.iloc[-1] / history.iloc[0]) - 1

    # MDD
    peak = history.cummax()
    drawdown = (history - peak) / peak
    mdd = drawdown.min()

    # Annualized Return
    days = (history.index[-1] - history.index[0]).days
    ann_return = (1 + total_return)**(365.25 / days) - 1

    # Calmar Ratio
    calmar = ann_return / abs(mdd) if mdd != 0 else 0

    return {
        'Total Return': total_return,
        'Annualized Return': ann_return,
        'Max Drawdown': mdd,
        'Calmar Ratio': calmar,
        'Drawdown Series': drawdown
    }

if __name__ == "__main__":
    df = load_and_clean_data('資料-1.1.xlsx')
    indicators = calculate_indicators(df)
    history = run_backtest(df, indicators)
    metrics = analyze_performance(history)

    print("\nBacktest Results:")
    for k, v in metrics.items():
        if k != 'Drawdown Series':
            print(f"{k}: {v:.4f}")

    mdd_date = metrics['Drawdown Series'].idxmin()
    print(f"Max Drawdown Date: {mdd_date}")

    # Identify regime of major loss (around MDD)
    # Let's see the context of MDD
    window = metrics['Drawdown Series'].loc[mdd_date - pd.Timedelta(days=60) : mdd_date]
    print("\nDrawdown around MDD:")
    print(window.tail(10))
