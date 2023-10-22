import pandas as pd


def save_data_to_excel(data, excel_path):
    # Use the number of columns in the first row as the expected number
    expected_columns = len(data[0])

    # Filter out rows with a different number of columns
    filtered_data = [row for row in data if len(row) == expected_columns]

    # Create DataFrame and save to Excel
    df = pd.DataFrame(filtered_data[1:], columns=filtered_data[0])
    df.to_excel(excel_path, index=False)
