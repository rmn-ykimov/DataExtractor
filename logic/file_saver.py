import pandas as pd


def save_data_to_excel(data, excel_path):
    """
    Save data to an Excel file.

    Parameters:
    data : list of lists of str
        The data to be saved, where each nested list represents a data row.
        excel_path : str
        The path to save the Excel file.

    Returns:
    None. The function's output is the creation of an Excel file.
    """

    # Determine the number of columns based on the first row of data
    expected_columns = len(data[0])

    # Filter rows to have the expected number of columns
    filtered_data = [row for row in data if len(row) == expected_columns]

    # Create a DataFrame based on the filtered data
    # The first row is used as column names
    df = pd.DataFrame(filtered_data[1:], columns=filtered_data[0])

    # Save the filtered DataFrame to an Excel file
    df.to_excel(excel_path, index=False)
