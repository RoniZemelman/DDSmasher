
import os
import time
from docx import Document  
import tkinter as tk
from tkinter.filedialog import askdirectory
from docx.shared import Cm

# Input Folder Path from Client through GUI 
root = tk.Tk()
root.withdraw() # keep the root window from appearing
folder_path = askdirectory() # show an "Open" dialog box in Explorer and return the path to the selected file

# TEST TIMING
start_time = time.time()

# Build corresponding lists of file, folder paths
# These lists will be used to determine the size of the table
# as well as to populate it
fileNames = []
filePaths = []
for root, dirs, files in os.walk(folder_path):

    for file in files:
        fileNames.append(file)
        filePaths.append(str(os.path.relpath(root, start = os.curdir)))

# **** WORD TABLE ****
document = Document()

style = document.styles['Normal']
font = style.font
font.bold= False

#changing the page margins
sections = document.sections
for section in sections:
    section.top_margin = Cm(0.5)
    section.bottom_margin = Cm(0.5)
    section.left_margin = Cm(0.5)
    section.right_margin = Cm(0.5)

# Create Word Doc Table
doc_table = document.add_table(rows = len(fileNames) + 2, cols=5)  # NOTE: KEY OPTIMIZING STEP
doc_table.style = 'Medium Shading 1 Accent 1'
doc_table.allow_autofit = True

# Adding column titles to the 1st row of the table
row = doc_table.rows[0].cells
row[0].text = 'Item #'
row[1].text = 'Folder Path'  
row[2].text = 'File Name'
row[3].text = "Summary"
row[4].text = "Flags"

doc_table.cell(0,0).width = 4846320    # 5.3 * 914400 
doc_table.cell(1,0).width = 4846320   

# NOTE: Key optimizing step 
# Read through all files, folders in corresponding lists
# then populate pre-constructed table.  
doc_table_cells = doc_table._cells  

index = 1
for item in fileNames:

     # WORD
    row_cells = doc_table_cells[index*5:(index+1)*5] # NOTE: Key optimizing Step
    row_cells[0].text = f"{index}." 
    row_cells[1].text = filePaths[index - 1]    # root 
    row_cells[2].text = fileNames[index - 1]
    index += 1

root_path = os.path.dirname(folder_path)
cleaned_folder_path = folder_path.replace(root_path, "")
cleaned_folder_path = cleaned_folder_path.replace("/", "")

document.save(f'{cleaned_folder_path}_FolderDD.docx')

print(f"It took {time.time() - start_time} seconds to execute {len(fileNames)} files for {folder_path}")
