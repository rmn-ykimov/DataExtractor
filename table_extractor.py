from docx import Document


def process_text(text):
    return text.replace('-\n', '').replace('\n', ' ')


def is_section_year_row(row):
    non_empty_cells = 0
    contains_year = False
    for cell in row:
        if cell:
            non_empty_cells += 1
            if "г." in cell:
                contains_year = True
    return non_empty_cells == 1 and contains_year


def extract_table_data(docx_path):
    document = Document(docx_path)
    table_data = []
    current_section_year = None
    headers_added = False

    for table in document.tables:
        current_table_data = []

        for row in table.rows:
            row_data = [process_text(cell.text.strip()) for cell in row.cells]

            if is_section_year_row(row_data):
                current_section_year = next(
                    (cell for cell in row_data if cell), None)
                continue

            row_data.append(
                "Год (Раздел)" if not headers_added else current_section_year)
            headers_added = True

            current_table_data.append(row_data)

        table_data.extend(current_table_data)

    return table_data
