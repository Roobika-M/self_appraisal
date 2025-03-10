from flask import Flask, request, render_template, redirect, url_for, make_response,flash,send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
import logging
import os
import pandas as pd
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
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

n=0
m = 0
# Fill the table3 with Excel data
###################table 1-journal###########################
for i, (_, row) in enumerate(df_filtered.iterrows()):
    row_index = start_row+i
    if i+2 >= len(table3.rows):  
        table3.add_row()  # Add rows if needed

    table3.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table3.cell(row_index+1, 1).text = str(row.get("Paper Title", "-"))
    table3.cell(row_index+1, 2).text = str(row.get("Journal Name", "-"))
    table3.cell(row_index+1, 3).text = str(row.get("Year of Publication", "-"))
    table3.cell(row_index+1, 4).text = str(row.get("ISSN", "-"))
    table3.cell(row_index+1, 5).text = str(row.get("Web Link", "-"))
    table3.cell(row_index+1, 6).text = str(row.get("Impact Factor", "-"))
    if row.get("Impact Factor", "-")!='-':
        if row.get("Impact Factor", "-")>3:
            n+=3
        elif row.get("Impact Factor", "-")>1.5 and row.get("Impact Factor", "-")<=3:
            n+=2
        elif row.get("Impact Factor", "-")>=1 and row.get("Impact Factor", "-")<=1.5:
            n+=1
    n+=2 

row_index=-1
table3.add_row()
table3.rows[row_index].cells[0].merge(table3.rows[row_index].cells[-2])
paragraph = table3.cell(-1, 5).paragraphs[0]
run = paragraph.add_run("Total score")
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
table3.cell(-1, 6).text = str(n)
m = m+n
#######################table2-book au#################################  
n = 0
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
    table4.cell(row_index+1, 1).text = str(row.get("Book Title", "-"))
    table4.cell(row_index+1, 2).text = str(row.get("Publication Name", "-"))
    table4.cell(row_index+1, 3).text = str(row.get("Date of Publication", "-"))  # Fixed Date Mapping
    table4.cell(row_index+1, 4).text = str(row.get("ISBN", "-"))
    table4.cell(row_index+1, 5).text = str(row.get("Description", "-"))  
################################table3-invited lectures##########################################
# no data shiiiiiiiiiiiiiiiiiiiiiiii
#######################################table4,5-conference##################################
n = 0
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
table = 0
for i, (_, row) in enumerate(df_filtered.iterrows()):
    conference_type = str(row["Conference Type"]).strip().lower()
    if conference_type == "international":
        table = table6
        organized_by = row.get("Organized By", "-")
        n += 2
    else:
        table = table7
        organized_by = row.get("Organized By", "-")
        n += 1
    row_index = start_row + i
    
    if i + 2 >= len(table.rows):
        table.add_row()  # Add rows if needed
    
    table.cell(row_index + 1, 0).text = str(i + 1)  # Serial No.
    table.cell(row_index + 1, 1).text = str(row.get("Paper Title", "-"))
    table.cell(row_index + 1, 2).text = organized_by
    table.cell(row_index + 1, 3).text = str(row.get("From Date", "-"))
    table.cell(row_index + 1, 4).text = str(row.get("Place", "-"))
    table.cell(row_index + 1, 5).text = str(row.get("Role", "-"))
if table!=0:
    row_index=-1
    table.add_row()
    table.rows[row_index].cells[0].merge(table.rows[row_index].cells[-2])
    paragraph = table.cell(-1, 4).paragraphs[0]
    run = paragraph.add_run("Total score")
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    table.cell(-1, 5).text = str(n)
m = m+n
##################################table6 research grant##############################################
n = 0
total_amt=0
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
    if str(row.get("Coordinator", "-"))=="applied":
        table8.cell(row_index+1, 0).text = str(i+1)  # Serial No.
        table8.cell(row_index+1, 1).text = str(row.get("Coordinator", "-"))
        table8.cell(row_index+1, 2).text = str(row.get("Title", "-")) 
        table8.cell(row_index+1, 3).text = str(row.get("Type", "-"))
        table8.cell(row_index+1, 4).text = str(row.get("Funding Agent", "-"))  
        table8.cell(row_index+1, 5).text = str(row.get("Amount", "-"))  
        table8.cell(row_index+1, 6).text = str(row.get("Applied On", "-"))  

        if row.get("Amount", "-")!="-":
            total_amt+=row.get("Amount", "-")
if row.get("Amount", "-")>1000000:
    n+=(row.get("Amount", "-")//1000000)*2
row_index=-1
table8.add_row()
table8.rows[row_index].cells[0].merge(table8.rows[row_index].cells[-2])
paragraph = table8.cell(-1, 5).paragraphs[0]
run = paragraph.add_run("Total score")
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
table8.cell(-1, 6).text = str(n)
m = m+n
#########################################table 7-seminar#########################################
n = 0
df_seminar = pd.read_excel(excel_path, sheet_name="Research Grant", skiprows=5)

df_seminar.columns = df_seminar.columns.str.strip()

df_seminar["Name of the faculty"] = df_seminar["Name of the faculty"].ffill()
df_filtered = df_seminar[df_seminar["Name of the faculty"].str.strip() == name]
table9_index = 9 
table9 = doc.tables[table9_index]
start_row = 1
if not df_filtered.empty:
    for i, (_, row) in enumerate(df_filtered.iterrows()):
        row_index = start_row+i
        if i+2 >= len(table9.rows):  
            table9.add_row()  # Add rows if needed
        if str(row.get("Coordinator", "-"))!="applied":
            table9.cell(row_index+1, 0).text = str(i+1)  # Serial No.
            table9.cell(row_index+1, 1).text = str(row.get("Coordinator", "-"))
            table9.cell(row_index+1, 2).text = str(row.get("Title", "-")) 
            table9.cell(row_index+1, 3).text = str(row.get("Type", "-"))
            table9.cell(row_index+1, 4).text = str(row.get("Funding Agent", "-"))  
            table9.cell(row_index+1, 5).text = str(row.get("Amount", "-"))  
            table9.cell(row_index+1, 6).text = str(row.get("Applied On", "-"))  

            if row.get("Amount", "-")!="-":
                total_amt+=row.get("Amount", "-")
    if row.get("Amount", "-")>50000:
        n+=(row.get("Amount", "-")//50000)
    row_index=-1
    table9.add_row()
    table9.rows[row_index].cells[0].merge(table9.rows[row_index].cells[-2])
    paragraph = table9.cell(-1, 5).paragraphs[0]
    run = paragraph.add_run("Total score")
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    table9.cell(-1, 6).text = str(n)
    m = m+n
###############################################table 8,9-patent##################
n = 0
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
    title = str(row.get("Title", "-")).strip()
    clear_and_write(table10.cell(row_index + 1, 1), title)

    # Determine Filing/Publishing Date
    status = str(row.get("Status", "")).strip().lower()
    date_value = str(row.get("Date", "-"))

    if status == "filed":
        clear_and_write(table10.cell(row_index + 1, 2), date_value)  # Date of Filing
        clear_and_write(table10.cell(row_index + 1, 3), "-")  # No publishing date
    elif status == "published":
        clear_and_write(table10.cell(row_index + 1, 2), "-")  # No filing date
        clear_and_write(table10.cell(row_index + 1, 3), date_value)  # Date of Publish
        n+=5
    else:
        clear_and_write(table10.cell(row_index + 1, 2), "-")
        clear_and_write(table10.cell(row_index + 1, 3), "-")

    # Fill Other Details
    clear_and_write(table10.cell(row_index + 1, 4), str(row.get("Status", "-")))

row_index=-1
table10.add_row()
table10.rows[row_index].cells[0].merge(table10.rows[row_index].cells[-2])
paragraph = table10.cell(-1, 3).paragraphs[0]
run = paragraph.add_run("Total score")
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
table10.cell(-1, 4).text = str(n)
m = m+n
research = m

#####################################table 10-consultancy####################
#table consultancy no data da
########################################table 11-workshop#################
m = 0
n = 0
df_workshop = pd.read_excel(excel_path, sheet_name="Workshops", skiprows=5)

df_workshop.columns = df_workshop.columns.str.strip()

df_workshop["Name of the faculty"] = df_workshop["Name of the faculty"].ffill()
df_filtered = df_workshop[df_workshop["Name of the faculty"].str.strip() == name]
table13_index = 14
table13 = doc.tables[table13_index]
start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    
    from_date = str(row.get("From Date","-"))
    to_date = str(row.get("To Date","-"))
    row_index = start_row+i
    if i+2 >= len(table13.rows):  
        table13.add_row()  # Add rows if needed
    if str(row.get("Role")) == "attended":
        table13.cell(row_index+1, 0).text = str(i+1)  # Serial No.
        table13.cell(row_index+1, 1).text = str(row.get("Topic", "-"))
        table13.cell(row_index+1, 2).text = f"{from_date} to {to_date}"
        table13.cell(row_index+1, 3).text = str(row.get("Description", "-"))  
        table13.cell(row_index+1, 4).text = str(row.get("Venue", "-")) 
        if n<3:
            n+=1

row_index=-1
table13.add_row()
table13.rows[row_index].cells[0].merge(table13.rows[row_index].cells[-2])
paragraph = table13.cell(-1, 3).paragraphs[0]
run = paragraph.add_run("Total score")
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
table13.cell(-1, 4).text = str(n)
m = m+n

###################################table 12-skill dev####################################
n = 0
df_develop = pd.read_excel(excel_path, sheet_name="Faculty Internship", skiprows=5)

df_develop.columns = df_develop.columns.str.strip()

df_develop["Name of the faculty"] = df_develop["Name of the faculty"].ffill()
df_filtered = df_develop[df_develop["Name of the faculty"].str.strip() == name]
table14_index = 15  
table14 = doc.tables[table14_index]
start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    from_date = str(row.get("From Date","-"))
    to_date = str(row.get("To Date","-"))
    row_index = start_row+i
    if i+2 >= len(table14.rows):  
        table14.add_row()  # Add rows if needed

    table14.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table14.cell(row_index+1, 1).text = str(row.get("FDP Name", "-"))
    table14.cell(row_index+1, 2).text = f"{from_date} to {to_date}"
    table14.cell(row_index+1, 3).text = str(row.get("Description", "-"))  
    table14.cell(row_index+1, 4).text = str(row.get("National or International", "-"))
    n += 3

row_index=-1
table14.add_row()
table14.rows[row_index].cells[0].merge(table14.rows[row_index].cells[-2])
paragraph = table14.cell(-1, 3).paragraphs[0]
run = paragraph.add_run("Total score")
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
table14.cell(-1, 4).text = str(n)
m = m+n
#############################table 13-mooc&nptel#############################################
n = 0
df_mooc = pd.read_excel(excel_path,sheet_name="MOOC Course", skiprows=4)
df_mooc.columns = df_mooc.columns.str.strip()
df_mooc["Name of the faculty"] = df_mooc["Name of the faculty"].ffill()
df_filtered = df_mooc[df_mooc["Name of the faculty"].str.strip() == name]
table15_index = 16
table15 = doc.tables[table15_index]
start_row = 1
for i, (_, row) in enumerate(df_filtered.iterrows()):
    from_date = str(row.get("From Date","-"))
    to_date = str(row.get("To Date","-"))
    row_index = start_row+i
    if i+2 >= len(table15.rows):  
        table15.add_row()  # Add rows if needed

    table15.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table15.cell(row_index+1, 1).text = str(row.get("Coure Title", "-"))
    table15.cell(row_index+1, 2).text = str(row.get("Course Type", "-"))
    table15.cell(row_index+1, 3).text = f"{from_date} to {to_date}"
    table15.cell(row_index+1, 4).text = str(row.get("Duration", "-")) 
    table15.cell(row_index+1, 5).text = str(row.get("Awards","-"))
    if n < 4:
        n += 2

row_index=-1
table15.add_row()
table15.rows[row_index].cells[0].merge(table15.rows[row_index].cells[-2])
paragraph = table15.cell(-1, 4).paragraphs[0]
run = paragraph.add_run("Total score")
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
table15.cell(-1, 5).text = str(n)
m = m+n
##################################table 14-mou##################################
# n = 0
# df_mou = pd.read_excel(excel_path,sheet_name="MoU", skiprows=2)
# df_mou.columns = df_mou.columns.str.strip()
# df_mou["Name of the faculty"] = df_mou["Name of the faculty"].ffill()
# df_filtered = df_mou[df_mou["Name of the faculty"].str.strip() == name]
# table16_index = 17
# table16 = doc.tables[table16_index]
# start_row = 1
# for i, (_, row) in enumerate(df_filtered.iterrows()):
#     from_date = str(row.get("From Date","-"))
#     to_date = str(row.get("To Date","-"))
#     row_index = start_row+i
#     if i+2 >= len(table16.rows):  
#         table16.add_row()  # Add rows if needed

#     table16.cell(row_index+1, 0).text = str(i+1)  # Serial No.
#     table16.cell(row_index+1, 1).text = str(row.get("Name of the faculty", "-"))
#     table16.cell(row_index+1, 2).text = str(row.get("Company Name", "-"))
#     table16.cell(row_index+1, 3).text = f"{from_date} to {to_date}"
#     table16.cell(row_index+1, 4).text = str(row.get("Industry SPOC", "-")) 
#     table16.cell(row_index+1, 5).text = str(row.get("Duration","-"))
#     n += 1

# if n>0:
#     table16.add_row()
#     table16.cell(row_index+2, 5).text = str(n)
#     m = m+n
###################################table 16-spl contribute##################################

# df_awards = pd.read_excel(excel_path,sheet_name="Extension Activities", skiprows=5)
# df_awards.columns = df_awards.columns.str.strip()
# df_awards["Name of the faculty"] = df_awards["Name of the faculty"].ffill()
# df_filtered = df_awards[df_awards["Name of the faculty"].str.strip() == name]
# table17_index = 18
# table17 = doc.tables[table17_index]
# start_row = 1
# for i, (_, row) in enumerate(df_filtered.iterrows()):
#     from_date = str(row.get("From Date","-"))
#     to_date = str(row.get("To Date","-"))
#     row_index = start_row+i
#     if i+2 >= len(table17.rows):  
#         table17.add_row()  # Add rows if needed

#     table17.cell(row_index+1, 0).text = str(i+1)  # Serial No.
#     table17.cell(row_index+1, 1).text = str(row.get("Name of the Event", "-"))
#     table17.cell(row_index+1, 2).text = f"{from_date} to {to_date}"
#     table17.cell(row_index+1, 3).text = str(row.get("Recognition", "-"))  
#     table17.cell(row_index+1, 4).text = str(row.get("Award", "-"))
#     table17.cell(row_index+1, 5).text = str(row.get("Description","-"))
###################################table 16-no of conference,workshop,hack###############################################
n = 0
df_workshop = pd.read_excel(excel_path, sheet_name="Workshops", skiprows=5)

df_workshop.columns = df_workshop.columns.str.strip()

df_workshop["Name of the faculty"] = df_workshop["Name of the faculty"].ffill()
df_filtered = df_workshop[df_workshop["Name of the faculty"].str.strip() == name]
table18_index = 19
table18 = doc.tables[table18_index]
start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    
    from_date = str(row.get("From Date","-"))
    to_date = str(row.get("To Date","-"))
    row_index = start_row+i
    if i+2 >= len(table18.rows):  
        table18.add_row()  # Add rows if needed
    if str(row.get("Role", "-"))=="conducted":
        table18.cell(row_index+1, 0).text = str(i+1)  # Serial No.
        table18.cell(row_index+1, 1).text = str(row.get("Topic", "-"))
        table18.cell(row_index+1, 2).text = str(row.get("Department", "-"))
        table18.cell(row_index+1, 3).text = f"{from_date} to {to_date}"
        table18.cell(row_index+1, 4).text = str(row.get("No of Students", "-"))  
        table18.cell(row_index+1, 5).text = str(row.get("Venue", "-"))  
        table18.cell(row_index+1, 6).text = str(row.get("Description", "-")) 

        n += 0.5

row_index=-1
table18.add_row()
table18.rows[row_index].cells[0].merge(table18.rows[row_index].cells[-2])
paragraph = table18.cell(-1, 5).paragraphs[0]
run = paragraph.add_run("Total score")
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
table18.cell(-1, 6).text = str(n)
m = m+n
################################################table 17-expert visit###############################
n = 0
df_experts = pd.read_excel(excel_path, sheet_name="Guest Lectures", skiprows=8)

df_experts.columns = df_experts.columns.str.strip()

df_experts["Faculty Name"] = df_experts["Faculty Name"].ffill()
df_filtered = df_experts[df_experts["Faculty Name"].str.strip() == name]
table19_index = 20
table19 = doc.tables[table19_index]
start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    
    from_date = str(row.get("From Date","-"))
    to_date = str(row.get("To Date","-"))
    row_index = start_row+i
    if i+2 >= len(table19.rows):  
        table19.add_row()  # Add rows if needed

    table19.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table19.cell(row_index+1, 1).text = str(row.get("Chief Guest Name", "-"))
    table19.cell(row_index+1, 2).text = str(row.get("Address", "-"))
    table19.cell(row_index+1, 3).text = str(row.get("Topic Name","-"))
    table19.cell(row_index+1, 4).text = f"{from_date} to {to_date}"  
    table19.cell(row_index+1, 5).text = str(row.get("Description", "-"))  
    table19.cell(row_index+1, 6).text = str(row.get("Topic Delivered", "-")) 
    n += 1


row_index=-1
table19.add_row()
table19.rows[row_index].cells[0].merge(table19.rows[row_index].cells[-2])
paragraph = table19.cell(-1, 5).paragraphs[0]
run = paragraph.add_run("Total score")
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
table19.cell(-1, 6).text = str(n)
m = m+n
selfm = m+n
#####################table 18-###########################################################
m = 0
n=0
df_project = pd.read_excel(excel_path, sheet_name="Project Guided and Mentoring")

df_project.columns = df_project.columns.str.strip()

df_project["Faculty Name"] = df_project["Faculty Name"].ffill()
df_filtered = df_project[df_project["Faculty Name"].str.strip() == name]
table21_index = 22
table21 = doc.tables[table21_index]
start_row = 1

for i, (_, row) in enumerate(df_filtered.iterrows()):
    
    from_date = str(row.get("From Date","-"))
    to_date = str(row.get("To Date","-"))
    row_index = start_row+i
    if i+2 >= len(table21.rows):  
        table21.add_row()  # Add rows if needed

    table21.cell(row_index+1, 0).text = str(i+1)  # Serial No.
    table21.cell(row_index+1, 1).text = str(row.get("Project Title", "-"))
    table21.cell(row_index+1, 2).text = str(row.get("Number of Students", "-"))
    table21.cell(row_index+1, 3).text = str(row.get("Title of Hackathon","-"))
    table21.cell(row_index+1, 4).text = str(row.get("Organized By", "-")) 
    table21.cell(row_index+1, 5).text = str(row.get("Date", "-"))  
    table21.cell(row_index+1, 6).text = str(row.get("Status", "-"))
    n=1

row_index=-1
table21.add_row()
table21.rows[row_index].cells[0].merge(table21.rows[row_index].cells[-2])
paragraph = table21.cell(-1, 5).paragraphs[0]
run = paragraph.add_run("Total score")
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
table21.cell(-1, 6).text = str(n)
m = m+n
mentor = m+n

placeholders = {
        "{{research}}": str(research),
        "{{self}}": str(selfm),
        "{{mentorship}}": str(mentor)
    }

# Replace placeholders in paragraphs with the corresponding values
for placeholder, value in placeholders.items():
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, value)
# Save the modified document
output_doc_path = "filled_template.docx"
doc.save(output_doc_path)
print(f"Word document saved as {output_doc_path}")