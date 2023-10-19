import pandas as pd


def save_data_to_excel(data, excel_path):
    df = pd.DataFrame(data[1:], columns=data[0])
    df.to_excel(excel_path, index=False)
