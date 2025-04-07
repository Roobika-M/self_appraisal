from flask import Flask, request, render_template, redirect, url_for, make_response,flash,send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
import logging
import os
import pandas as pd
from docx import Document
from docx2pdf import convert
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from docx.shared import Pt
import pythoncom
import win32com.client
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

app = Flask(__name__)
# app.secret_key = "your_secret_key"

# app.config['SQLALCHEMY_DATABASE_URI'] = 'postgresql://postgres:harshu8564@localhost/KITE_STAFF'
# app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# db = SQLAlchemy(app)

# class userlo(db.Model):
#     id = db.Column(db.Integer, primary_key=True)
#     usernam = db.Column(db.String(80), nullable=False, unique=True)
#     password = db.Column(db.String(120), nullable=False)

#     def __repr__(self):
#         return f"<User {self.usernam}>"

# class staff(db.Model):
#     id = db.Column(db.Integer, primary_key=True)
#     name = db.Column(db.String(100), nullable=False)
#     position = db.Column(db.String(50))

#     def __repr__(self):
#         return f"<Staff {self.name}>"

# Render the login page
@app.route('/')
def home():
    return render_template('login.html', error=None)

# Handle form submission
@app.route('/login', methods=['POST', 'GET'])
def login():
    if request.method == 'POST':  
        username = request.form.get('username')
        password = request.form.get('password')

        if not username or not password:
            error = "Please enter both username and password."
            return render_template('login.html', error=error)

        # usr = userlo.query.filter_by(usernam=username).first()
        usr=username

        # if usr and password == usr.password:
        if usr and password == "admin":
            return redirect(url_for('upload'))
        else:
            error = "Please enter correct username and password."
            return render_template('login.html', error=error)

    return render_template('login.html')
staffname = ''
detaillist = []
excel_path = ''

@app.route('/upload', methods=['POST', 'GET'])
def upload():
    global staffname, detaillist, excel_path

    if request.method == 'POST':
        print("🚀 Form submitted!")  # Debugging
        
        # Get form data
        name = request.form.get('name')
        designation = request.form.get('designation')
        department = request.form.get('dept')
        emp_id = request.form.get('empid')

        # Validate form data
        if not all([name, designation, department, emp_id]):
            print("⚠️ Missing form data!")  # Debugging
            return render_template('upload.html', error="Please fill in all details.")

        staffname = name
        detaillist = [name, designation, department, emp_id]

        # Get uploaded file
        file = request.files.get('excel_file')
        if not file or file.filename == '':
            print("⚠️ No file uploaded!")  # Debugging
            return render_template('upload.html', error="Please upload an Excel file.")

        # Save the file
        upload_folder = os.getcwd()
        file_path = os.path.join(upload_folder, file.filename)
        file.save(file_path)

        excel_path = file.filename
        print(f"📂 File saved at: {file_path}")  # Debugging

        try:
            processing(excel_path, staffname)
        except Exception as e:
            print(f"❌ Error processing file: {e}")  # Debugging
            return render_template('upload.html', error="Error processing file.")

        print("✅ Redirecting to download page...")  # Debugging
        return redirect(url_for("download_path"))

    return render_template('upload.html')
    

@app.route('/data', methods=['POST', 'GET'])
def data():
    if request.method == 'POST':
        id = request.form.get('id')
        username = request.form.get('username')
        name = request.form.get('name')
        password = request.form.get('password')
        position = request.form.get('posi')

        # if not all([username, name, password, position]):
        #     error = "Please enter all details."
        #     return render_template('data.html', error=error)
        # existing_user = userlo.query.filter_by(usernam=username).first()

        # if existing_user:
        #     error = "Username exist."
        #     return render_template('data.html', error=error)
        
        # try:
        #     new_user = userlo(id=id,usernam=username, password=password)
        #     new_staff = staff(id=id,name=name, position=position)

        #     db.session.add(new_user)
        #     db.session.add(new_staff)
        #     db.session.commit()

        #     flash("User added successfully!", "success")
        #     return redirect(url_for('login'))
        # except Exception as e:
        #     db.session.rollback()
        #     flash(f"Error: {str(e)}", "error")
        #     return redirect(url_for('data'))

    return render_template('data.html')

'''logging.basicConfig()
logging.getLogger('sqlalchemy.engine').setLevel(logging.INFO)'''

@app.route('/download/<file_type>', methods=['GET'])
def download(file_type):
    upload_folder = os.getcwd()  # Get current working directory
    docx_filename = os.path.join(upload_folder, "filled_template.docx")
    pdf_filename = os.path.join(upload_folder, "filled_template.pdf")

    if file_type == "docx":
        if not os.path.exists(docx_filename):
            return "File not found", 404
        return send_file(docx_filename, as_attachment=True)

    elif file_type == "pdf":
        convert_docx_to_pdf(docx_filename, pdf_filename)
        if not os.path.exists(pdf_filename):
            return "PDF conversion failed", 500
        return send_file(pdf_filename, as_attachment=True)

    return "Invalid file type", 400

@app.route('/download_path')
def download_path():
    global staffname, research, selfm
    total_score = research + selfm
    return render_template("download.html", name=staffname, research=research, selfm=selfm, total_score=total_score)


def convert_docx_to_pdf(docx_path, pdf_path):
    pythoncom.CoInitialize()  # Fixes COM error
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Run in background

    try:
        docx_path = os.path.abspath(docx_path)  # Ensure absolute path
        pdf_path = os.path.abspath(pdf_path)  # Ensure absolute path
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # Convert DOCX to PDF
        doc.Close()
    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        word.Quit()  # Close Word
        pythoncom.CoUninitialize()  # Clean up COM

#################################### Load the Excel file
# Load the Word document
def processing(excel_path,staffname):
    doc_path = "template.docx"
    doc = Document(doc_path)
    name = staffname
    global m, n, research,selfm  # Declare m and n as global variables
    m = 0
    n = 0
    research = 0
    selfm = 0
######################################################################################################
    df_journal = pd.read_excel(excel_path, sheet_name="Journal Publication", skiprows=5)

    # Fix column names (remove spaces)
    df_journal.columns = df_journal.columns.str.strip()
    df_filtered=[]
    df_journal["Name of the faculty"] = df_journal["Name of the faculty"].ffill()
    df_filtered = df_journal[df_journal["Name of the faculty"].str.strip() == name]
    m = 0
    if not df_filtered.empty:
        table3_index = 3  
        table3 = doc.tables[table3_index]

        start_row = 1

        n=0
        
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
    df_filtered=[]
    df_bookpub["Name of the faculty"] = df_bookpub["Name of the faculty"].ffill()
    df_filtered = df_bookpub[df_bookpub["Name of the faculty"].str.strip() == name]
    if not df_filtered.empty:
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
    df_filtered=[]
    df_conference.columns = df_conference.columns.str.strip()
    df_conference["Name of the faculty"] = df_conference["Name of the faculty"].ffill()

    df_filtered = df_conference[df_conference["Name of the faculty"].str.strip() == name]
    if not df_filtered.empty:
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
    df_filtered=[]
    df_research = pd.read_excel(excel_path, sheet_name="Research Grant", skiprows=5)

    df_research.columns = df_research.columns.str.strip()

    df_research["Name of the faculty"] = df_research["Name of the faculty"].ffill()
    df_filtered = df_research[df_research["Name of the faculty"].str.strip() == name]
    if not df_filtered.empty:
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
    df_filtered=[]
    df_seminar.columns = df_seminar.columns.str.strip()

    df_seminar["Name of the faculty"] = df_seminar["Name of the faculty"].ffill()
    df_filtered = df_seminar[df_seminar["Name of the faculty"].str.strip() == name]
    if not df_filtered.empty:
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
    df_filtered=[]
    df_patent.columns = df_patent.columns.str.strip()

    df_patent["Faculty name"] = df_patent["Faculty name"].ffill()
    df_filtered = df_patent[df_patent["Faculty name"].str.strip() == name]
    if not df_filtered.empty:
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
    df_filtered=[]
    df_workshop["Name of the faculty"] = df_workshop["Name of the faculty"].ffill()
    df_filtered = df_workshop[df_workshop["Name of the faculty"].str.strip() == name]
    if not df_filtered.empty:
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
    df_filtered=[]
    df_develop = pd.read_excel(excel_path, sheet_name="Faculty Internship", skiprows=5)
    df_develop.columns = df_develop.columns.str.strip()
    possible_names=["Name of the faculty","Faculty Name"] 
    selectedcol=next((col for col in possible_names if col in df_develop.columns),None)
    if selectedcol:
        df_develop[selectedcol] = df_develop[selectedcol].ffill()
        df_filtered = df_develop[df_develop[selectedcol].str.strip() == name]
    if not df_filtered.empty:
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
    df_filtered=[]
    df_mooc = pd.read_excel(excel_path,sheet_name="MOOC Course", skiprows=4)
    df_mooc.columns = df_mooc.columns.str.strip()
    possible_names=["Name of the faculty","Faculty Name"] 
    selectedcol=next((col for col in possible_names if col in df_mooc.columns),None)
    if selectedcol:
        df_mooc[selectedcol] = df_mooc[selectedcol].ffill()
        df_filtered = df_mooc[df_mooc[selectedcol].str.strip() == name]
    if not df_filtered.empty: 
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
    n = 0
    df_filtered=[]
    df_mou = pd.read_excel(excel_path,sheet_name="MoU", skiprows=5)
    df_mou.columns = df_mou.columns.str.strip()
    possible_names=["Name of the faculty","Faculty Name"] 
    selectedcol=next((col for col in possible_names if col in df_mou.columns),None)
    if selectedcol:
        df_mou[selectedcol] = df_mou[selectedcol].ffill()
        df_filtered = df_mou[df_mou[selectedcol].str.strip() == name]
    if not df_filtered.empty:
        table16_index = 17
        table16 = doc.tables[table16_index]
        start_row = 1
        for i, (_, row) in enumerate(df_filtered.iterrows()):
            from_date = str(row.get("From Date","-"))
            to_date = str(row.get("To Date","-"))
            row_index = start_row+i
            if i+2 >= len(table16.rows):  
                table16.add_row()  # Add rows if needed

            table16.cell(row_index+1, 0).text = str(i+1)  # Serial No.
            table16.cell(row_index+1, 1).text = str(row.get(selectedcol, "-"))
            table16.cell(row_index+1, 2).text = str(row.get("Company Name", "-"))
            table16.cell(row_index+1, 3).text = f"{from_date} to {to_date}"
            table16.cell(row_index+1, 4).text = str(row.get("Industry SPOC", "-")) 
            table16.cell(row_index+1, 5).text = str(row.get("Duration","-"))
            n += 1

        row_index=-1
        table16.add_row()
        table16.rows[row_index].cells[0].merge(table14.rows[row_index].cells[-2])
        paragraph = table16.cell(-1, 4).paragraphs[0]
        run = paragraph.add_run("Total score")
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        table16.cell(-1, 5).text = str(n)
        m = m+n
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
    df_filtered=[]
    df_workshop.columns = df_workshop.columns.str.strip()

    df_workshop["Name of the faculty"] = df_workshop["Name of the faculty"].ffill()
    df_filtered = df_workshop[(df_workshop["Name of the faculty"].str.strip() == name) & (df_workshop["Role"].fillna("").str.strip() == "conducted")]
    if not df_filtered.empty:
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
    df_filtered=[]
    df_experts = pd.read_excel(excel_path, sheet_name="Guest Lectures", skiprows=8)

    df_experts.columns = df_experts.columns.str.strip()

    df_experts["Faculty Name"] = df_experts["Faculty Name"].ffill()
    df_filtered = df_experts[df_experts["Faculty Name"].str.strip() == name]
    if not df_filtered.empty:
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
    df_filtered=[]
    df_project = pd.read_excel(excel_path, sheet_name="Project Guided and Mentoring")

    df_project.columns = df_project.columns.str.strip()

    df_project["Faculty Name"] = df_project["Faculty Name"].ffill()
    df_filtered = df_project[df_project["Faculty Name"].str.strip() == name]
    if not df_filtered.empty:
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
            "{{mentorship}}": str(mentor),
            "{{name}}":detaillist[0],
            "{{designation}}":detaillist[1],
            "{{dept}}":detaillist[2],
            "{{empid}}":detaillist[3]
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

if __name__ == '__main__':
    # with app.app_context():
    #     db.create_all()
    app.run(debug=True)