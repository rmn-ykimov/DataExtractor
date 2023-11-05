import re
from docx import Document


def process_text(text):
    """
    Process text in a cell, removing line breaks and replacing them with spaces.

    Parameters:
    text : str
        The text to be processed.

    Returns:
    str
        The processed text.
    """
    return text.replace('-\n', '').replace('\n', ' ')


def is_section_year_row(row):
    """
    Check whether the given row is a row with a year (or section).

    Parameters:
    row : list of str
        The list of cells in a row.

    Returns:
    bool
        True if the row contains only one non-empty cell with a year, otherwise False.
    """

    # Count the number of non-empty cells in the row
    non_empty_cells = sum(1 for cell in row if cell.strip())

    # Compile a regular expression to search for a year pattern in the cell
    year_regex = re.compile(r'^\s*(\d{4})(\s+год|\s+г\.|\s*года)?\s*$')

    # Check if at least one cell in the row contains a year matching the regular expression
    contains_year = any(year_regex.match(cell) for cell in row)

    # Return True if exactly one cell is non-empty and contains a year
    return non_empty_cells == 1 and contains_year


def extract_table_data(docx_path):
    """
    Extract data from tables in a DOCX file.

    Parameters:
    docx_path : str
        The path to the DOCX file to be processed.

    Returns:
    list of lists of str
        The data from all tables in the file.
    """

    # Open the DOCX file
    document = Document(docx_path)

    # Initialize variables
    table_data = []  # Data for all tables
    current_section_year = None  # Current year or section
    headers_added = False  # Flag for adding the "Год (Раздел)" header

    # Iterate through all tables in the document
    for table in document.tables:
        current_table_data = []  # Data for the current table

        # Iterate through all rows in the table
        for row in table.rows:
            # Process each cell in the row
            row_data = [process_text(cell.text.strip()) for cell in row.cells]

            # Check for a row with a year (or section)
            if is_section_year_row(row_data):
                current_section_year = next(
                    (cell for cell in row_data if cell), None)
                continue

            # Add the year (or section) to the row
            row_data.append(
                "Год (Раздел)" if not headers_added else current_section_year)
            headers_added = True

            # Save the processed row
            current_table_data.append(row_data)

        # Add the data for the current table to the overall list
        table_data.extend(current_table_data)

    # Return the data for all tables
    return table_data
