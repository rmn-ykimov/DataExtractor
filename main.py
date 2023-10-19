import os
from table_extractor import extract_table_data
from file_saver import save_data_to_excel


def main():
    current_dir = os.getcwd()
    docx_path = os.path.join(current_dir, "input.docx")
    extracted_data = extract_table_data(docx_path)

    output_excel_path = os.path.join(current_dir, "output.xlsx")
    save_data_to_excel(extracted_data, output_excel_path)

    print(f"Path for excel file is {output_excel_path}")


if __name__ == "__main__":
    main()
