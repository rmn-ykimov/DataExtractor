import pandas as pd
from config import unwanted_values


def save_data_to_excel(data, excel_path):

    # Determine the number of columns based on the first row of data
    expected_columns = len(data[0])

    # Filter rows to have the expected number of columns
    filtered_data = [row for row in data if len(row) == expected_columns]

    # Create a DataFrame based on the filtered data
    # The first row is used as column names
    df = pd.DataFrame(filtered_data[1:], columns=filtered_data[0])

    # Filter the DataFrame to remove rows with unwanted values
    for column in df.columns:
        df = df[df[column].apply(lambda x: x not in unwanted_values)]

    # Save the filtered DataFrame to an Excel file
    df.to_excel(excel_path, index=False)
