import pandas as pd
import numpy as np
import os

def tag_etf(row):
    name = str(row['Name'])
    symbol = str(row['Symbol'])
    tags = []

    # 1. Leveraged / Inverse (Exclusive categories for primary type, but separated here)
    if any(k in name for k in ['正2', '正２']):
        tags.append('槓桿型')
    if any(k in name for k in ['反1', '反１']):
        tags.append('反向型')

    # 2. Regional Equity Tags
    is_equity = False
    if any(k in name for k in ['中國', '滬', '深', '上証', '上證', '恒生', '香港', 'A股', '中証', '亞太', '韓']):
        tags.append('中國及亞太股票')
        is_equity = True
    if any(k in name for k in ['美國', 'S&P', '費城', '那斯達克', 'NASDAQ', '道瓊', '標普', 'MAG7', '北美', 'FANG+', '洲際半導體', 'ARK']):
        tags.append('美國股票')
        is_equity = True
    if any(k in name for k in ['日本', '日經', '東證']):
        tags.append('日本股票')
        is_equity = True
    if any(k in name for k in ['越南', '印度', '歐洲', '全球', '新興', '太空', '潔淨能源', 'AI', '網路資安', '機器人', '基因', '電池', '儲能', '元宇宙', '數位支付']):
        tags.append('其他/全球股票')
        is_equity = True

    # Taiwan Equity Sub-categorization
    is_taiwan = False
    if any(k in name for k in ['台', '臺灣', '加權', '櫃', '中型100', '藍籌', '市值', '旗艦', '50']):
        if '美國' not in name and '中國' not in name and '日本' not in name:
            tags.append('台灣股票')
            is_taiwan = True
            is_equity = True

    if is_taiwan or (len(symbol) == 4 and symbol.isdigit() and not symbol.endswith('B')):
        if '台灣股票' not in tags: tags.append('台灣股票')
        is_taiwan = True
        is_equity = True
        if any(k in name for k in ['電子', '半導體', '科技', 'IC設計', '通訊', '5G']):
            tags.append('台股-電子')
        if any(k in name for k in ['金融', '銀行']):
            tags.append('台股-金融')
        if any(k in name for k in ['生技', '醫療', '基因']):
            tags.append('台股-生技')
        if any(k in name for k in ['高息', '股息', '填息', '優息']):
            tags.append('台股-高股息')
        if any(k in name for k in ['ESG', '永續', '低碳']):
            tags.append('台股-ESG/永續')
        if any(k in name for k in ['加權', '市值', '50', '100', '加權']):
            tags.append('台股-指數型')

    # 3. Bonds Sub-categorization
    if symbol.endswith('B') or any(k in name for k in ['債', '入息', '收益', '投等']):
        tags.append('債券/固定收益')
        if any(k in name for k in ['公債', '美債', '國債']):
            tags.append('債券-公債')
        if any(k in name for k in ['公司債', '投等債', '非投等', '信用債', 'A級']):
            tags.append('債券-公司債')
        if any(k in name for k in ['金融債', '銀行債']):
            tags.append('債券-金融債')
        if any(k in name for k in ['1-3', '0-1', '短期']):
            tags.append('債券-短期')
        if any(k in name for k in ['7-10', '10年']):
            tags.append('債券-中長期')
        if any(k in name for k in ['20年', '25年', '長債']):
            tags.append('債券-長期')

    # 4. Commodities
    if any(k in name for k in ['黃金', '石油', '原油', '黃豆', '白銀', '銅', '航運', '原物料']):
        tags.append('商品型')

    # 5. Currency
    if any(k in name for k in ['美元', '日圓', '匯率']):
        tags.append('匯率型')

    # 6. Real Estate
    if any(k in name for k in ['地產', 'REITs', '不動產']):
        tags.append('房地產')

    if not tags:
        tags.append('其他/未分類')

    return tags

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

    # 2. Tagging (Multi-label)
    print("Tagging ETFs...")
    etf_info['Tags'] = etf_info.apply(tag_etf, axis=1)

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

    # 4. Merge and Analyze Tags
    data = pd.merge(etf_info, returns_df, on='Symbol')

    # Explode tags to analyze each tag separately
    exploded_data = data.explode('Tags')
    tag_stats = exploded_data.groupby('Tags').agg({
        'TotalReturn': 'mean',
        'Symbol': 'count'
    }).rename(columns={'Symbol': 'Count'}).sort_values(by='TotalReturn', ascending=False)

    # 5. Generate Report
    print("Generating report...")
    report_content = f"""# ETF 深度分類與多維度表現分析報告

## 1. 類別績效概覽 (含細分標籤)
數據週期：{cleaned_dates[0]} 至 {cleaned_dates[-1]}
本分析將 ETF 進行多重標籤劃分，同一檔 ETF 可歸類於多個細分類別（例如：某 ETF 可同時屬於「台灣股票」與「台股-高股息」）。

### 各標籤績效排名：
{tag_stats.to_markdown()}

---

## 2. 強勢類別深度分析

"""
    # Analyze Top 5 tags
    top_tags = tag_stats.index[:5]
    for i, tag in enumerate(top_tags, 1):
        tag_data = exploded_data[exploded_data['Tags'] == tag].sort_values(by='TotalReturn', ascending=False)
        report_content += f"### **第{i}名：{tag}**\n"
        report_content += f"*   平均收益率：{tag_stats.loc[tag, 'TotalReturn']:.2%}\n"
        report_content += f"*   該類別成分股數量：{tag_stats.loc[tag, 'Count']}\n"
        report_content += f"*   最強成份股 (前5名)：\n"
        report_content += tag_data[['Symbol', 'Name', 'TotalReturn']].head(5).to_markdown(index=False)
        report_content += "\n\n"

    report_content += """## 3. 策略分析與建議

### **策略可行性 (多維度觀察)**
1.  **交叉驗證強勢訊號：** 當多個相關標籤（如「台灣股票」與「台股-電子」）同時位居前列時，代表該趨勢具有極高的可信度。
2.  **精準標的選取：** 透過細分標籤，您可以從廣泛的「債券」類別中，精確發現是「公債」強還是「公司債」強，從而做出更細膩的配置。
3.  **風險對沖參考：** 獨立觀察「槓桿型」與「反向型」，可以清楚看到目前市場的情緒波動與對沖成本。

### **實務建議**
1.  **聚焦子產業領頭羊：** 在強勢大類別中，尋找表現最突出的子標籤（如：台股中的半導體或高股息），並挑選其前三名成份股。
2.  **債券配置轉向：** 觀察不同存續期間（長債 vs 短債）的表現差異，在利率變動週期中動態調整。
3.  **警惕標籤重疊：** 由於同一 ETF 可能有多個標籤，配置時需注意是否過度集中於某一特定產業或性質。

---

## 4. 完整 ETF 清單與標籤對應
以下為所有 ETF 的完整列表及其所屬的所有標籤：

"""
    # Sort for final listing
    final_list = data.sort_values(by='Symbol')
    final_list['Tags_Str'] = final_list['Tags'].apply(lambda x: ', '.join(x))
    final_list_display = final_list[['Symbol', 'Name', 'TotalReturn', 'Tags_Str']].copy()
    final_list_display['TotalReturn'] = final_list_display['TotalReturn'].map(lambda x: f"{x:.2%}" if pd.notnull(x) else "N/A")

    report_content += final_list_display.to_markdown(index=False)

    report_content += """\n\n---
*備註：以上數據基於「ETF報價分類.xlsx」計算。*
"""
    with open('ETF分析報告.md', 'w', encoding='utf-8') as f:
        f.write(report_content)
    print("Analysis complete. Report saved to ETF分析報告.md")

if __name__ == "__main__":
    main()
