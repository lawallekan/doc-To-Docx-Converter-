import os
import win32com.client as win32

# Source folder containing .doc files
source_folder = "C:\\Users\\RAZER\\Downloads\\forex"

# Destination folder to save .docx files
destination_folder = "C:\\Users\\RAZER\\Downloads\\forex\\docx"

# Create destination folder if it doesn't exist
if not os.path.exists(destination_folder):
    os.makedirs(destination_folder)

# List all files in the source folder
files = [f for f in os.listdir(source_folder) if f.endswith('.doc')]

# Initialize Microsoft Word
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False

# Loop through each .doc file and save it as .docx
for file in files:
    doc_path = os.path.join(source_folder, file)
    docx_path = os.path.join(destination_folder, file + 'x')
    
    # Open .doc file
    doc = word.Documents.Open(doc_path)
    
    # Save as .docx
    doc.SaveAs(docx_path, FileFormat=16)
    
    # Close the document
    doc.Close()

# Quit Microsoft Word
word.Quit()