import os


# The function to create a directory, if it is not exist
def create_directory_if_not_exists(directory_path):
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)


# The function for getting a list of all DOCX files in a specified directory
def get_all_docx_files_in_directory(directory_path):
    docx_files = []
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith('.docx'):
                docx_path = os.path.join(root, file)
                docx_files.append(docx_path)
    return docx_files
