from docx import Document


def process_text(text):
    text = text.replace('-\n', '').replace('\n', ' ')
    return text


def extract_table_data(docx_path):
    document = Document(docx_path)
    table_data = []

    for table in document.tables:
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                processed_text = process_text(cell.text.strip())
                row_data.append(processed_text)
            table_data.append(row_data)

    return table_data
