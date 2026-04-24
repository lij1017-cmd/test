import pandas as pd
import numpy as np
import os

def categorize_etf(row):
    name = str(row['Name'])
    symbol = str(row['Symbol'])

    if any(k in name for k in ['正2', '正２', '反1', '反１']):
        return '杠桿/反向'

    if symbol.endswith('B') or any(k in name for k in ['債', '入息', '收益', '投等', '金融股息']):
        return '債券/固定收益'

    if any(k in name for k in ['黃金', '石油', '原油', '黃豆', '白銀', '銅', '航運', '原物料']):
        return '商品型'

    if any(k in name for k in ['美元', '日圓', '匯率']):
        return '匯率型'

    if any(k in name for k in ['地產', 'REITs', '不動產']):
        return '房地產'

    if any(k in name for k in ['中國', '滬', '深', '上証', '上證', '恒生', '香港', 'A股', '中証', '亞太', '恒生', '韓']):
        return '中國及亞太股票'

    if any(k in name for k in ['美國', 'S&P', '費城', '那斯達克', 'NASDAQ', '道瓊', '標普', 'MAG7', '北美', 'FANG+', '洲際半導體', 'ARK']):
        return '美國股票'

    if any(k in name for k in ['日本', '日經', '東證']):
        return '日本股票'

    if any(k in name for k in ['越南', '印度', '歐洲', '全球', '新興', '太空', '潔淨能源', 'AI', '網路資安', '機器人', '基因', '電池', '儲能', '元宇宙', '數位支付']):
        return '其他/全球股票'

    if any(k in name for k in ['台', '臺灣', '加權', '櫃', '中型100', '科技', '電子', '金融', '股息', '高息', 'ESG', '市值', '小資', '精選', '優息', '存股', '龍頭', '永續', '智慧', '低碳', '價值', 'IC設計', '半導體', '電動車', '未來車', '智能車', '低波', '藍籌', '公司治理', '工業30', '優選', '旗艦', '數據', '平衡', '動態', '未來50', '填息', '強棒']):
        return '台灣股票'

    if len(symbol) == 4 and symbol.isdigit():
        return '台灣股票'

    return '其他/未分類'

def main():
    file_path = 'ETF報價分類.xlsx'
    if not os.path.exists(file_path):
        print(f"Error: {file_path} not found.")
        return

    # 1. Data Extraction
    print("Extracting data...")
    df_prices = pd.read_excel(file_path, sheet_name='收盤價', header=None)
    symbols = [str(s) for s in df_prices.iloc[0, 5:].values]
    names = df_prices.iloc[1, 5:].values
    dates_raw = df_prices.iloc[3:, 4].values
    prices_raw = df_prices.iloc[3:, 5:].values
    cleaned_dates = [str(d)[:8] for d in dates_raw]

    price_df = pd.DataFrame(prices_raw, index=cleaned_dates, columns=symbols)
    price_df = price_df.apply(lambda x: pd.to_numeric(x, errors='coerce'))

    etf_info = pd.DataFrame({'Symbol': symbols, 'Name': names})

    # 2. Categorization
    print("Categorizing ETFs...")
    etf_info['Category'] = etf_info.apply(categorize_etf, axis=1)

    # 3. Calculate Returns
    print("Calculating returns...")
    returns = {}
    for col in price_df.columns:
        series = price_df[col].dropna()
        if len(series) >= 2:
            returns[col] = (series.iloc[-1] / series.iloc[0]) - 1
        else:
            returns[col] = np.nan
    returns_df = pd.DataFrame(list(returns.items()), columns=['Symbol', 'TotalReturn'])

    # 4. Merge and Analyze
    data = pd.merge(etf_info, returns_df, on='Symbol')
    category_stats = data.groupby('Category').agg({
        'TotalReturn': 'mean',
        'Symbol': 'count'
    }).rename(columns={'Symbol': 'Count'}).sort_values(by='TotalReturn', ascending=False)

    # 5. Generate Report
    print("Generating report...")
    report_content = f"""# ETF 分類與表現分析報告

## 1. ETF 分類概覽
數據週期：{cleaned_dates[0]} 至 {cleaned_dates[-1]}
本報告將 {len(etf_info)} 檔 ETF 劃分為 {len(category_stats)} 個類別。

### 類別績效排名：
{category_stats.to_markdown()}

---

## 2. 強勢類別與領先成份股分析

"""
    top_categories = category_stats.index[:3]
    for i, cat in enumerate(top_categories, 1):
        cat_data = data[data['Category'] == cat].sort_values(by='TotalReturn', ascending=False)
        report_content += f"### **第{i}名：{cat}**\n"
        report_content += f"*   平均收益率：{category_stats.loc[cat, 'TotalReturn']:.2%}\n"
        report_content += f"*   最強成份股：\n"
        report_content += cat_data[['Symbol', 'Name', 'TotalReturn']].head(5).to_markdown(index=False)
        report_content += "\n\n"

    report_content += """## 3. 可行性分析與建議

### **可行性分析**
您的策略屬於典型的 **「由上而下」(Top-Down)** 的動能投資策略。這種方法在金融實務中具有高度可行性：
1.  **趨勢追隨 (Trend Following)：** 類別指數的強勢通常代表了宏觀經濟趨勢（如能源漲價或技術革新），這種趨勢往往具有持續性。
2.  **分散風險：** 透過觀察類別而非單一股票，可以過濾掉個別公司的隨機風險。
3.  **效率配置：** 資源集中在最強勢的市場軌道，提高資金效率。

### **策略建議**
1.  **風險調整：** 收益最高的類別（如商品）往往波動也大。建議參考夏普比率（收益/風險）進行篩選。
2.  **時機配合：** 結合技術指標（如類別指數是否站上均線）來決定進場點。
3.  **定期再平衡：** 市場強弱會輪動。建議每季重新執行此分析，確保配置在最強勢的類別。
4.  **注意槓桿損耗：** 槓桿/反向型 ETF 存在每日調整損耗，不建議作為長期持有標的。

---

## 4. 全部分類清單 (附收益率)
以下為各分類下的完整 ETF 表列：

"""
    # List all categories and their ETFs
    for cat in category_stats.index:
        cat_data = data[data['Category'] == cat].sort_values(by='Symbol')
        report_content += f"### **【{cat}】 ({len(cat_data)}檔)**\n"
        # Optional: Format return as percentage for readability in the final table
        cat_data_display = cat_data[['Symbol', 'Name', 'TotalReturn']].copy()
        cat_data_display['TotalReturn'] = cat_data_display['TotalReturn'].map(lambda x: f"{x:.2%}" if pd.notnull(x) else "N/A")
        report_content += cat_data_display.to_markdown(index=False)
        report_content += "\n\n"

    report_content += """---
*備註：以上數據基於「ETF報價分類.xlsx」計算。*
"""
    with open('ETF分析報告.md', 'w', encoding='utf-8') as f:
        f.write(report_content)
    print("Analysis complete. Report saved to ETF分析報告.md")

if __name__ == "__main__":
    main()
