import os
from logic.table_extractor import extract_table_data
from logic.file_saver import save_data_to_excel


# Function to process a single DOCX file and save its tables to an Excel file
def process_file(docx_path):
    # Extract tables from the DOCX file
    extracted_data = extract_table_data(docx_path)

    # Create the name for the output Excel file based on the DOCX file name
    base_name = os.path.basename(docx_path).replace('.docx', '.xlsx')

    # Determine the full path for the output Excel file
    output_excel_path = os.path.join(os.getcwd(), 'output', base_name)

    # Save the extracted data to an Excel file
    save_data_to_excel(extracted_data, output_excel_path)

    # Print the path where the Excel file has been saved
    print(f"Path for excel file is {output_excel_path}")
