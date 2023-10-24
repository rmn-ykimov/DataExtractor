from docx import Document


def process_text(text):
    return text.replace('-\n', '').replace('\n', ' ')


def is_section_year_row(row):
    non_empty_cells = sum(1 for cell in row if cell)
    return non_empty_cells == 1 and any("г." in cell_content for cell_content in row)


def extract_table_data(docx_path):
    document = Document(docx_path)
    table_data = []
    current_section_year = None
    headers_added = False

    for table in document.tables:
        for row in table.rows:
            row_data = [process_text(cell.text.strip()) for cell in row.cells]

            if is_section_year_row(row_data):
                current_section_year = next(
                    (cell for cell in row_data if cell), None)
                continue

            if not headers_added:
                row_data.append("Год (Раздел)")
                headers_added = True

            else:
                row_data.append(current_section_year)

            table_data.append(row_data)

    return table_data
