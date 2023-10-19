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
