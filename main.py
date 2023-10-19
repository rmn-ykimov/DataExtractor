import os
import pandas as pd

from docx import Document


def extract_table_data(docx_path):
    document = Document(docx_path)

    table_data = []

    for table in document.tables:
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text.strip())
            table_data.append(row_data)

    return table_data


current_dir = os.getcwd()

docx_path = os.path.join(current_dir, "input.docx")

extracted_data = extract_table_data(docx_path)

print(extracted_data)

df = pd.DataFrame(extracted_data[1:], columns=extracted_data[0])

output_excel_path = os.path.join(current_dir, "output.xlsx")
df.to_excel(output_excel_path, index=False)

print(output_excel_path)
