import pandas as pd


def save_data_to_excel(data, excel_path):
    expected_columns = len(data[0])
    filtered_data = [row for row in data if len(row) == expected_columns]
    df = pd.DataFrame(filtered_data[1:], columns=filtered_data[0])
    df.to_excel(excel_path, index=False)
