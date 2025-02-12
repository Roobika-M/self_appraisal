import pandas as pd
from docx import Document
from docx.shared import Pt

# Load the Excel file
excel_path = "data.xlsx"
# Load the Word document
doc_path = "template.docx"
doc = Document(doc_path)
name = input("Enter the name:")
######################################################################################################
df_journal = pd.read_excel(excel_path, sheet_name="Journal Publication", skiprows=5)

# Fix column names (remove spaces)
df_journal.columns = df_journal.columns.str.strip()

df_journal["Name of the faculty"] = df_journal["Name of the faculty"].ffill()
df_filtered = df_journal[df_journal["Name of the faculty"].str.strip() == name]

table3_index = 3  
table3 = doc.tables[table3_index]

start_row = 1
# Fill the table3 with Excel data
for i, (_, row) in enumerate(df_filtered.iterrows()):
    row_index = start_row+i
    if i+2 >= len(table3.rows):  
        table3.add_row()  # Add rows if needed

    table3.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table3.cell(row_index+1, 1).text = str(row.get("Paper Title", "N/A"))
    table3.cell(row_index+1, 2).text = str(row.get("Journal Name", "N/A"))
    table3.cell(row_index+1, 3).text = str(row.get("Year of Publication", "N/A"))
    table3.cell(row_index+1, 4).text = str(row.get("ISSN", "N/A"))
    table3.cell(row_index+1, 5).text = str(row.get("Citation", "N/A"))
    table3.cell(row_index+1, 6).text = str(row.get("Impact Factor", "N/A"))
#############################################################################
df_bookpub = pd.read_excel(excel_path, sheet_name="Book Publication", skiprows=5)

df_bookpub.columns = df_bookpub.columns.str.strip()

df_bookpub["Name of the faculty"] = df_bookpub["Name of the faculty"].ffill()
df_filtered = df_bookpub[df_bookpub["Name of the faculty"].str.strip() == name]
table4_index = 4  
table4 = doc.tables[table4_index]
start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    row_index = start_row+i
    if i+2 >= len(table4.rows):  
        table4.add_row()  # Add rows if needed

    table4.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table4.cell(row_index+1, 1).text = str(row.get("Book Title", "N/A"))
    table4.cell(row_index+1, 2).text = str(row.get("Publication Name", "N/A"))
    table4.cell(row_index+1, 3).text = str(row.get("Name of the faculty", "N/A"))  # Fixed Date Mapping
    table4.cell(row_index+1, 4).text = str(row.get("ISBN", "N/A"))
    table4.cell(row_index+1, 5).text = str(row.get("Description", "N/A"))  
##########################################################################
#table 5 no data shiiiiiiiiiiiiiiiiiiiiiiii
##########################################################################

df_conference = pd.read_excel(excel_path, sheet_name="Conferences", skiprows=6)

df_conference.columns = df_conference.columns.str.strip()
df_conference["Name of the faculty"] = df_conference["Name of the faculty"].ffill()

df_filtered = df_conference[df_conference["Name of the faculty"].str.strip() == name]

# Get table references
table6_index = 6  # International conference table
table7_index = 7  # National conference table
table6 = doc.tables[table6_index]
table7 = doc.tables[table7_index]

start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    conference_type = str(row["Conference Type"]).strip().lower()
    if conference_type == "international":
        table = table6
        organized_by = row.get("Name of the International Conference", "N/A")
    else:
        table = table7
        organized_by = row.get("Name of the national Conference", "N/A")
    
    row_index = start_row + i
    
    if i + 2 >= len(table.rows):
        table.add_row()  # Add rows if needed
    
    table.cell(row_index + 1, 0).text = str(i + 1)  # Serial No.
    table.cell(row_index + 1, 1).text = str(row.get("Paper Title", "N/A"))
    table.cell(row_index + 1, 2).text = organized_by
    table.cell(row_index + 1, 3).text = str(row.get("From Date", "N/A"))
    table.cell(row_index + 1, 4).text = str(row.get("ISSN", "N/A"))
    table.cell(row_index + 1, 5).text = str(row.get("Citation", "N/A"))
################################################################################

df_research = pd.read_excel(excel_path, sheet_name="Research Grant", skiprows=5)

df_research.columns = df_research.columns.str.strip()

df_research["Name of the faculty"] = df_research["Name of the faculty"].ffill()
df_filtered = df_research[df_research["Name of the faculty"].str.strip() == name]
table8_index = 8  
table8 = doc.tables[table8_index]
start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    row_index = start_row+i
    if i+2 >= len(table8.rows):  
        table8.add_row()  # Add rows if needed

    table8.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table8.cell(row_index+1, 1).text = str(row.get("Principal", "N/A"))
    table8.cell(row_index+1, 2).text = str(row.get("Title", "N/A"))
    table8.cell(row_index+1, 3).text = str(row.get("Department", "N/A"))  
    table8.cell(row_index+1, 4).text = str(row.get("Type", "N/A"))
    table8.cell(row_index+1, 5).text = str(row.get("Funding Agent", "N/A"))  
    table8.cell(row_index+1, 6).text = str(row.get("Amount", "N/A"))  
    table8.cell(row_index+1, 7).text = str(row.get("Year", "N/A"))  
    table8.cell(row_index+1, 8).text = str(row.get("Duration", "N/A"))  

##################################################################################

df_seminar = pd.read_excel(excel_path, sheet_name="Seminar", skiprows=4)

df_seminar.columns = df_seminar.columns.str.strip()

df_seminar["Name of the faculty"] = df_seminar["Name of the faculty"].ffill()
df_filtered = df_seminar[df_seminar["Name of the faculty"].str.strip() == name]
table9_index = 9 
table9 = doc.tables[table9_index]
start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    row_index = start_row+i
    if i+2 >= len(table9.rows):  
        table9.add_row()  # Add rows if needed

    table9.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table9.cell(row_index+1, 1).text = str(row.get("Co-ordinator", "N/A"))
    table9.cell(row_index+1, 2).text = str(row.get("Types", "N/A"))
    table9.cell(row_index+1, 3).text = str(row.get("Type", "N/A"))  
    table9.cell(row_index+1, 4).text = str(row.get("Sponsored By", "N/A"))
    table9.cell(row_index+1, 5).text = str(row.get("Amount", "N/A"))  
    table9.cell(row_index+1, 6).text = str(row.get("Year", "N/A"))
    table9.cell(row_index+1, 6).text = str(row.get("Duration", "N/A"))

#################################################################

df_patent = pd.read_excel(excel_path, sheet_name="Patents")

df_patent.columns = df_patent.columns.str.strip()

df_patent["Faculty name"] = df_patent["Faculty name"].ffill()
df_filtered = df_patent[df_patent["Faculty name"].str.strip() == name]
table10_index = 10
table10 = doc.tables[table10_index]

start_row = 1

from docx.shared import Pt

def clear_and_write(cell, text):
    """ Clears the cell and writes new text with formatting. """
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.clear()  # Ensure old text is removed
    cell.text = ""  # Extra clearing step
    run = cell.paragraphs[0].add_run(text)
    run.font.size = Pt(11)  # Ensure text is visible
    run.font.bold = True  # Make text bold to check visibility

for i, (_, row) in enumerate(df_filtered.iterrows()):
    row_index = start_row + i
    if row_index + 1 >= len(table10.rows):  
        table10.add_row()  # Add new rows if needed

    # Debugging: Print before writing
    print(f"Writing to row {row_index+1}: {row.to_dict()}")

    # Fill Serial No.
    clear_and_write(table10.cell(row_index + 1, 0), str(i + 1))

    # Fill Patent Title
    title = str(row.get("Title", "N/A")).strip()
    clear_and_write(table10.cell(row_index + 1, 1), title)

    # Determine Filing/Publishing Date
    status = str(row.get("Status", "")).strip().lower()
    date_value = str(row.get("Date", "N/A"))

    if status == "filed":
        clear_and_write(table10.cell(row_index + 1, 2), date_value)  # Date of Filing
        clear_and_write(table10.cell(row_index + 1, 3), "-")  # No publishing date
    elif status == "published":
        clear_and_write(table10.cell(row_index + 1, 2), "-")  # No filing date
        clear_and_write(table10.cell(row_index + 1, 3), date_value)  # Date of Publish
    else:
        clear_and_write(table10.cell(row_index + 1, 2), "-")
        clear_and_write(table10.cell(row_index + 1, 3), "-")

    # Fill Other Details
    clear_and_write(table10.cell(row_index + 1, 4), str(row.get("Patent No", "N/A")))
    clear_and_write(table10.cell(row_index + 1, 5), str(row.get("Sponsored By", "N/A")))

#########################################################
#table consultancy no data da
#########################################################

df_workshop = pd.read_excel(excel_path, sheet_name="Workshops", skiprows=5)

df_workshop.columns = df_workshop.columns.str.strip()

df_workshop["Name of the faculty"] = df_workshop["Name of the faculty"].ffill()
df_filtered = df_workshop[df_workshop["Name of the faculty"].str.strip() == name]
table13_index = 13
table13 = doc.tables[table13_index]
start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    
    from_date = str(row.get("From Date","N/A"))
    to_date = str(row.get("To Date","N/A"))
    row_index = start_row+i
    if i+2 >= len(table13.rows):  
        table13.add_row()  # Add rows if needed

    table13.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table13.cell(row_index+1, 1).text = str(row.get("Topic", "N/A"))
    table13.cell(row_index+1, 2).text = f"{from_date} to {to_date}"
    table13.cell(row_index+1, 3).text = str(row.get("Description", "N/A"))  
    table13.cell(row_index+1, 4).text = str(row.get("Department", "N/A")) 
#######################################################################

df_develop = pd.read_excel(excel_path, sheet_name="Faculty Development Program", skiprows=5)

df_develop.columns = df_develop.columns.str.strip()

df_develop["Name of the Faculty"] = df_develop["Name of the Faculty"].ffill()
df_filtered = df_develop[df_develop["Name of the Faculty"].str.strip() == name]
table14_index = 14  
table14 = doc.tables[table14_index]
start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    from_date = str(row.get("From Date","N/A"))
    to_date = str(row.get("To Date","N/A"))
    row_index = start_row+i
    if i+2 >= len(table14.rows):  
        table14.add_row()  # Add rows if needed

    table14.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table14.cell(row_index+1, 1).text = str(row.get("FDP Name", "N/A"))
    table14.cell(row_index+1, 2).text = f"{from_date} to {to_date}"
    table14.cell(row_index+1, 3).text = str(row.get("Description", "N/A"))  
    table14.cell(row_index+1, 4).text = str(row.get("Department", "N/A"))
##########################################################################

df_mooc = pd.read_excel(excel_path,sheet_name="MOOC Course", skiprows=4)
df_mooc.columns = df_mooc.columns.str.strip()
df_mooc["Name of the faculty"] = df_mooc["Name of the faculty"].ffill()
df_filtered = df_mooc[df_mooc["Name of the faculty"].str.strip() == name]
table15_index = 15 
table15 = doc.tables[table15_index]
start_row = 1
for i, (_, row) in enumerate(df_filtered.iterrows()):
    from_date = str(row.get("From Date","N/A"))
    to_date = str(row.get("To Date","N/A"))
    row_index = start_row+i
    if i+2 >= len(table15.rows):  
        table15.add_row()  # Add rows if needed

    table15.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table15.cell(row_index+1, 1).text = str(row.get("Coure Title", "N/A"))
    table15.cell(row_index+1, 2).text = f"{from_date} to {to_date}"
    table15.cell(row_index+1, 3).text = str(row.get("Details", "N/A"))  
    table15.cell(row_index+1, 4).text = str(row.get("Department", "N/A"))
    table15.cell(row_index+1, 5).text = str(row.get("Awards","N/A"))
####################################################################
#######MoU data not found 16
#####################################################################

df_awards = pd.read_excel(excel_path,sheet_name="Extension Activities", skiprows=5)
df_awards.columns = df_awards.columns.str.strip()
df_awards["Name of the faculty"] = df_awards["Name of the faculty"].ffill()
df_filtered = df_awards[df_awards["Name of the faculty"].str.strip() == name]
table17_index = 17
table17 = doc.tables[table17_index]
start_row = 1
for i, (_, row) in enumerate(df_filtered.iterrows()):
    from_date = str(row.get("From Date","N/A"))
    to_date = str(row.get("To Date","N/A"))
    row_index = start_row+i
    if i+2 >= len(table17.rows):  
        table17.add_row()  # Add rows if needed

    table17.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table17.cell(row_index+1, 1).text = str(row.get("Name of the Event", "N/A"))
    table17.cell(row_index+1, 2).text = f"{from_date} to {to_date}"
    table17.cell(row_index+1, 3).text = str(row.get("Recognition", "N/A"))  
    table17.cell(row_index+1, 4).text = str(row.get("Award", "N/A"))
    table17.cell(row_index+1, 5).text = str(row.get("Description","N/A"))
##################################################################################
df_workshop = pd.read_excel(excel_path, sheet_name="Workshops", skiprows=5)

df_workshop.columns = df_workshop.columns.str.strip()

df_workshop["Name of the faculty"] = df_workshop["Name of the faculty"].ffill()
df_filtered = df_workshop[df_workshop["Name of the faculty"].str.strip() == name]
table18_index = 18
table18 = doc.tables[table18_index]
start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    
    from_date = str(row.get("From Date","N/A"))
    to_date = str(row.get("To Date","N/A"))
    row_index = start_row+i
    if i+2 >= len(table18.rows):  
        table18.add_row()  # Add rows if needed

    table18.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table18.cell(row_index+1, 1).text = str(row.get("Topic", "N/A"))
    table18.cell(row_index+1, 2).text = str(row.get("Department", "N/A"))
    table18.cell(row_index+1, 3).text = f"{from_date} to {to_date}"
    table18.cell(row_index+1, 4).text = str(row.get("No of Students", "N/A"))  
    table18.cell(row_index+1, 5).text = str(row.get("Venue", "N/A"))  
    table18.cell(row_index+1, 6).text = str(row.get("Description", "N/A")) 
###############################################################################
df_experts = pd.read_excel(excel_path, sheet_name="Guest Lectures", skiprows=8)

df_experts.columns = df_experts.columns.str.strip()

df_experts["Faculty Name"] = df_experts["Faculty Name"].ffill()
df_filtered = df_experts[df_experts["Faculty Name"].str.strip() == name]
table19_index = 19
table19 = doc.tables[table19_index]
start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    
    from_date = str(row.get("From Date","N/A"))
    to_date = str(row.get("To Date","N/A"))
    row_index = start_row+i
    if i+2 >= len(table19.rows):  
        table19.add_row()  # Add rows if needed

    table19.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table19.cell(row_index+1, 1).text = str(row.get("Chief Guest Name", "N/A"))
    table19.cell(row_index+1, 2).text = str(row.get("Address", "N/A"))
    table19.cell(row_index+1, 3).text = str(row.get("Topic Name","N/A"))
    table19.cell(row_index+1, 4).text = f"{from_date} to {to_date}"  
    table19.cell(row_index+1, 5).text = str(row.get("Description", "N/A"))  
    table19.cell(row_index+1, 6).text = str(row.get("Topic Delivered", "N/A")) 
################################################################################
df_project = pd.read_excel(excel_path, sheet_name="Project Guided and Mentoring")

df_project.columns = df_project.columns.str.strip()

df_project["Faculty Name"] = df_project["Faculty Name"].ffill()
df_filtered = df_project[df_project["Faculty Name"].str.strip() == name]
table21_index = 21
table21 = doc.tables[table21_index]
start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    
    from_date = str(row.get("From Date","N/A"))
    to_date = str(row.get("To Date","N/A"))
    row_index = start_row+i
    if i+2 >= len(table21.rows):  
        table21.add_row()  # Add rows if needed

    table21.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table21.cell(row_index+1, 1).text = str(row.get("Project Title", "N/A"))
    table21.cell(row_index+1, 2).text = str(row.get("Number of Students", "N/A"))
    table21.cell(row_index+1, 3).text = str(row.get("Thrust area","N/A"))
    table21.cell(row_index+1, 4).text = str(row.get("Outcome of the project", "N/A")) 
    table21.cell(row_index+1, 5).text = str(row.get("Interdisciplinary", "N/A"))  
    table21.cell(row_index+1, 6).text = str(row.get("Status", "N/A"))

# Save the modified document
output_doc_path = "filled_template.docx"
doc.save(output_doc_path)
print(f"Word document saved as {output_doc_path}")