from flask import Flask, request, render_template, redirect, url_for, make_response,flash,send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
import logging
import os
import pandas as pd
from docx import Document
from docx.shared import Pt

app = Flask(__name__)
app.secret_key = "your_secret_key"

app.config['SQLALCHEMY_DATABASE_URI'] = 'postgresql://postgres:harshu8564@localhost/KITE_STAFF'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

class userlo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    usernam = db.Column(db.String(80), nullable=False, unique=True)
    password = db.Column(db.String(120), nullable=False)

    def __repr__(self):
        return f"<User {self.usernam}>"

class staff(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    position = db.Column(db.String(50))

    def __repr__(self):
        return f"<Staff {self.name}>"

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

        usr = userlo.query.filter_by(usernam=username).first()

        if usr and password == usr.password:
            return redirect(url_for('upload'))
        else:
            error = "Please enter correct username and password."
            return render_template('login.html', error=error)

    return render_template('login.html')
staffname=''
# Upload page
@app.route('/upload', methods=['POST','GET'])
def upload():
    if request.method == 'POST':  # Corrected handling of form submission
        name = request.form.get('name')
        designation = request.form.get('designation')
        department = request.form.get('dept')
        emp_id = request.form.get('empid')

        global staffname
        staffname=name
        if not all([name, designation, department, emp_id]):
            error = "Please enter correct username and password."
            return render_template('upload.html', error=error)

        return redirect(url_for('excel'))  

    return render_template('upload.html')
    

@app.route('/data', methods=['POST', 'GET'])
def data():
    if request.method == 'POST':
        id = request.form.get('id')
        username = request.form.get('username')
        name = request.form.get('name')
        password = request.form.get('password')
        position = request.form.get('posi')

        if not all([username, name, password, position]):
            error = "Please enter all details."
            return render_template('data.html', error=error)
        existing_user = userlo.query.filter_by(usernam=username).first()

        if existing_user:
            error = "Username exist."
            return render_template('data.html', error=error)
        
        try:
            new_user = userlo(id=id,usernam=username, password=password)
            new_staff = staff(id=id,name=name, position=position)

            db.session.add(new_user)
            db.session.add(new_staff)
            db.session.commit()

            flash("User added successfully!", "success")
            return redirect(url_for('login'))
        except Exception as e:
            db.session.rollback()
            flash(f"Error: {str(e)}", "error")
            return redirect(url_for('data'))

    return render_template('data.html')

'''logging.basicConfig()
logging.getLogger('sqlalchemy.engine').setLevel(logging.INFO)'''

excel_path=''
@app.route('/excel', methods=['POST', 'GET'])
def excel():
    if request.method == 'POST':  # Fixed handling (was incorrectly checking GET)
        file = request.files.get('excel_file')

        if not file or file.filename == '':
            error = "Please upload the file."
            return render_template('excel.html', error=error)

        upload_folder = os.getcwd()
        app.config['UPLOAD_FOLDER'] = upload_folder
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)

        global staffname

        name=staffname
        global excel_path
        excel_path=file.filename
        processing(excel_path,name)

        flash(f"File {file.filename} successfully uploaded!", "success")
        return redirect(url_for("download_path"))

    return render_template('excel.html')



@app.route('/download', methods=['POST', 'GET'])
def download():
    filename=os.path.join(app.config['UPLOAD_FOLDER'],"filled_template.docx")
    return send_file(filename,as_attachment=True)

@app.route('/download_path')
def download_path():
    return render_template("download.html")



#################################### Load the Excel file
# Load the Word document
def processing(excel_path,staffname):
    doc_path = "template.docx"
    doc = Document(doc_path)
    name = staffname
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

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)


