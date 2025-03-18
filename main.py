import os
import win32com.client

def get_total_editing_time(folder_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Keep Word hidden
    
    total_time = 0  # Store total editing time in minutes
    
    for file in os.listdir(folder_path):
        if file.endswith(".doc") or file.endswith(".docx"):  # Check for Word files
            file_path = os.path.join(folder_path, file)
            try:
                doc = word.Documents.Open(file_path)  # Open the document
                editing_time = doc.BuiltInDocumentProperties("Total Editing Time").Value  # Get editing time
                total_time += editing_time  # Sum up total time
                print(f"{file}: {editing_time} minutes")
                doc.Close(False)  # Close without saving
            except Exception as e:
                print(f"Error reading {file}: {e}")
    
    word.Quit()
    print(f"\nTotal Editing Time for all documents: {total_time // 60} hr {total_time % 60} minutes")

# Set the folder path where your Word files are located
folder_path = r"The path to folder where Word files are located"
get_total_editing_time(folder_path)
