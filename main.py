import os
from file_processor import process_file
from file_utils import (create_directory_if_not_exists,
                        get_all_docx_files_in_directory)


# Main function to process all DOCX files in the 'input' directory
def main():

    # Determine the 'input' directory path
    input_dir = os.path.join(os.getcwd(), 'input')

    # Create the 'output' directory if it doesn't exist
    create_directory_if_not_exists(
        'output')

    # Loop through all files in the 'input' directory
    docx_files = get_all_docx_files_in_directory(
        input_dir)
    for docx_path in docx_files:
        process_file(docx_path)


# Entry point of the script
if __name__ == "__main__":
    main()
