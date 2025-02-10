from flask import Flask, request, render_template, redirect, url_for, make_response
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)

app.config['SQLALCHEMY_DATABASE_URI'] = 'postgresql://postgres:*****@localhost/KITE_STAFF'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

class userlogin(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user = db.Column(db.String(80), nullable=False, unique=True)
    password = db.Column(db.String(120), nullable=False)

    def __repr__(self):
        return f"<User {self.user}>"

class Staff(db.Model):
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
@app.route('/login', methods=['POST'])
@app.route('/login', methods=['POST'])
def login():
    username = request.form.get('username')
    password = request.form.get('password')

    if not username or not password:
        error = "Please enter both username and password."
        return render_template('login.html', error=error)

    usr = userlogin.query.filter_by(user=username).first()

    # Validate password and user existence
    if usr and password == usr.password:
        return redirect(url_for('upload'))
    else:
        error = "Invalid username or password. Please try again."
        response = make_response(render_template('login.html', error=error))
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        return response


# Upload page
@app.route('/upload', methods=['POST','GET'])
def upload():
    name = request.form.get('name')
    des = request.form.get('designation')
    dept = request.form.get('dept')
    empid = request.form.get('empid')
    return render_template('upload.html', error=None)

@app.route('/data', methods=['POST', 'GET'])
def data():
    if request.method == 'POST':
        id = request.form.get('id')
        username = request.form.get('username')
        name = request.form.get('name')
        password = request.form.get('password')
        position = request.form.get('posi')

        # Type conversion for ID
        if id:
            try:
                id = int(id)
            except ValueError:
                return "ID must be a number"

        if all([id, username, name, password, position]):
            try:
                user1 = userlogin(user=username, password=password, id=id)
                staff1 = Staff(name=name, position=position, id=id)

                db.session.add(user1)
                db.session.add(staff1)
                db.session.commit()
                return "Data successfully added!"
            except Exception as e:
                db.session.rollback()
                return f"Error: {str(e)}"
        else:
            error = "Please enter all values"
            response = make_response(render_template('data.html', error=error))
            response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
            return response
    return render_template('data.html')

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)
