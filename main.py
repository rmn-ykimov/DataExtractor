import os
from table_extractor import extract_table_data
from file_saver import save_data_to_excel


def process_file(docx_path):
    extracted_data = extract_table_data(docx_path)
    base_name = os.path.basename(docx_path).replace('.docx', '.xlsx')
    output_excel_path = os.path.join(os.getcwd(), 'output', base_name)
    save_data_to_excel(extracted_data, output_excel_path)
    print(f"Path for excel file is {output_excel_path}")


def main():
    input_dir = os.path.join(os.getcwd(), 'input')
    if not os.path.exists('output'):
        os.makedirs('output')

    for root, dirs, files in os.walk(input_dir):
        for file in files:
            if file.endswith('.docx'):
                docx_path = os.path.join(root, file)
                process_file(docx_path)


if __name__ == "__main__":
    main()
