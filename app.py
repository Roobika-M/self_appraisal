from flask import Flask, request, render_template, redirect, url_for, make_response,flash,send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
import logging
import os

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

# Upload page
@app.route('/upload', methods=['POST','GET'])
def upload():
    if request.method == 'POST':  # Corrected handling of form submission
        name = request.form.get('name')
        designation = request.form.get('designation')
        department = request.form.get('dept')
        emp_id = request.form.get('empid')

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

        flash(f"File {file.filename} successfully uploaded!", "success")
        return redirect(url_for("download_path"))

    return render_template('excel.html')

@app.route('/download', methods=['POST', 'GET'])
def download():
    filename=os.path.join(app.config['UPLOAD_FOLDER'],"KGCMS FINALLL WORD.docx")
    return send_file(filename,as_attachment=True)

@app.route('/download_path')
def download_path():
    return render_template("download.html")

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)
