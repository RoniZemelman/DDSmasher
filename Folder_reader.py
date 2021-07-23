
import os
import time
from docx import Document # Not recognized in IDE for some reason, but yes in command line
import tkinter as tk
from tkinter.filedialog import askdirectory
from prettytable import PrettyTable
from prettytable import MSWORD_FRIENDLY
from docx.shared import Inches, Cm

# TODO - Insert while loop that accepts and writes title for each folder


root = tk.Tk()
root.withdraw() # keep the root window from appearing
folder_path = askdirectory() # show an "Open" dialog box in Explorer and return the path to the selected file

# Create Table
table = PrettyTable()
table.set_style(MSWORD_FRIENDLY)
# table.field_names = ["Folder Path", "File Name"]

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

# Word Doc Table
doc_table = document.add_table(rows=2, cols=5)
doc_table.style = 'Medium Shading 1 Accent 1'
doc_table.allow_autofit = True

# Adding heading in the 1st row of the table
row = doc_table.rows[0].cells
row[0].text = 'Item #'
row[1].text = 'Folder Path'  
row[2].text = 'File Name'
row[3].text = "Summary"
row[4].text = "Flags"

doc_table.cell(0,0).width = 4846320    # 5.3 * 914400 
doc_table.cell(1,0).width = 4846320   


# TEST TIMING
start_time = time.time()

i = 1
numOfFiles = 0
numOfFolders = 0

for root, dirs, files in os.walk(folder_path):

    for dir in dirs:
        numOfFolders += 1

    for file in files:
        numOfFiles += 1
 
        # WORD
        row = doc_table.add_row().cells
        row[0].text = f"{i}." 
        row[1].text = os.path.relpath(root, start = os.curdir)    # root 
        row[2].text = file
        
        i += 1

# Word Doc
document.save('test.docx')

print(f"It took {time.time() - start_time} seconds to execute program of {numOfFiles} files and {numOfFolders} folders")
