import pandas as pd
import numpy as np

def clean_sheet_data(df):
    """
    Apply cleaning logic:
    - First valid data point back-fills leading gaps.
    - Middle gaps are forward-filled.
    """
    # Assuming the first column is the date and the first two rows are headers
    header_rows = df.iloc[:2, :].copy()
    data = df.iloc[2:, 1:].copy()

    # Convert data to numeric to handle NaN correctly
    data = data.apply(pd.to_numeric, errors='coerce')

    cleaned_data = data.copy()
    for col in cleaned_data.columns:
        series = cleaned_data[col]
        first_valid_idx = series.first_valid_index()
        if first_valid_idx is not None:
            # Back-fill leading NaNs from the first valid value
            val = series.loc[first_valid_idx]
            series.iloc[:series.index.get_loc(first_valid_idx)] = val
            # Forward-fill middle and trailing NaNs
            series = series.ffill()
        cleaned_data[col] = series

    result_df = df.copy()
    result_df.iloc[2:, 1:] = cleaned_data
    return result_df

def main():
    input_file = '資料-1.xlsx'
    output_file = '資料-1.1.xlsx'

    print(f"Cleaning data from {input_file}...")
    xl = pd.ExcelFile(input_file)
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name in xl.sheet_names:
            print(f"  Processing sheet: {sheet_name}")
            df = xl.parse(sheet_name, header=None)
            cleaned_df = clean_sheet_data(df)
            cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    print(f"Data cleaning complete. Result saved to {output_file}")

if __name__ == "__main__":
    main()
