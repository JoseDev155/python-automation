import os
from docxtpl import DocxTemplate
from datetime import datetime
import pandas as pd
import openpyxl

# Return generated files with format
def generated_files(n):
    return os.path.join("docs", f"generated_doc_{n}.docx")

# Dirs and base-generated documents
template_file = os.path.join("templates", "template.docx")
data_file = os.path.join("data", "data.xlsx")

doc = DocxTemplate(template_file)

# Const data
my_name = "Name"
date = datetime.today().strftime("%d %b, %Y")

# First dictionary
context = {'nombre' : my_name, 'fecha' : date}

# Read and extract from excel file
df = pd.read_excel(data_file, sheet_name="Hoja1")
for index, row in df.iterrows():
    # Second dictionary
    data_context = {'remitente' : row['name'], 'email_remitente' : row['email']}
    # Join dictionaries
    data_context.update(context)
    
    # Write and save data in the new files
    doc.render(data_context)
    generated_file = generated_files(index)
    doc.save(generated_file)