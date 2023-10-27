import pandas as pd


def save_data_to_excel(data, excel_path):
    expected_columns = len(data[0])
    filtered_data = [row for row in data if len(row) == expected_columns]

    df = pd.DataFrame(filtered_data[1:], columns=filtered_data[0])

    unwanted_values = [
        "№ п/п",
        "Делопроизводственные индексы или номера по старой описи",
        "Наименование единиц хранения",
        "Дата",
        "Количество листов",
        "Примечания",
        "1",
        "2",
        "3",
        "4",
        "5",
        "6"
    ]

    for column in df.columns:
        df = df[df[column].apply(lambda x: x not in unwanted_values)]

    df.to_excel(excel_path, index=False)
