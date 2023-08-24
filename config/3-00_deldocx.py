import os

def delete_non_temp_docx_files(directory):
    try:
        for root, _, files in os.walk(directory):
            for filename in files:
                if filename.endswith(".docx") and filename != "temp.docx":
                    file_path = os.path.join(root, filename)
                    os.remove(file_path)
                    print(f"Deleted: {file_path}")
        print("Deletion completed.")
    except OSError as e:
        print(f"Error while deleting files: {e}")

if __name__ == "__main__":
    template_directory = "./template"
    delete_non_temp_docx_files(template_directory)
