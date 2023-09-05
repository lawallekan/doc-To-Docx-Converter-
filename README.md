# doc-To-Docx-Converter-
code for  converting bulk  Doc in a folder to docx

Bulk .DOC to .DOCX Converter
This Python script automatically converts all Microsoft Word .doc files in a specified source folder to .docx format and saves them in a designated destination folder.

Requirements
Python 3.x
Microsoft Word installed on your system
Python libraries: python-docx and pywin32
Installation
First, install the required Python libraries:

bash
Copy code
pip install python-docx pywin32
Usage
Clone this repository or download the Python script.
Open the script with a text editor and set the source_folder and destination_folder variables to your specific folders.
source_folder: The folder containing your .doc files.
destination_folder: The folder where you want to save the .docx files.
Save the script.
Run the script.
bash
Copy code
python your_script_name.py
Replace your_script_name.py with the name you gave to the Python script.

Example
Here's a snippet of what the code looks like:

python
Copy code
# Source folder containing .doc files
source_folder = "C:\\Users\\Downloads\\forex"

# Destination folder to save .docx files
destination_folder = "C:\\Users\\Downloads\\forex\\docx"

# ... (rest of the code)
After running the script, all .doc files in the source_folder will be converted to .docx format and saved in the destination_folder.
