import pandas as pd
from flask import Flask, render_template, request, redirect, url_for,session,send_from_directory
import os
from openpyxl import Workbook, load_workbook
app = Flask(__name__)
app.secret_key = "your_secret_key"

UPLOAD_FOLDER = "uploads"
EXCEL_FILE = "submissions.xlsx"

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["EXCEL_FILE"] = EXCEL_FILE


@app.route('/')
def index():
    return render_template('index.html')

@app.route("/students")
def students():
    return render_template("students.html")


@app.route("/student_submit", methods=["POST"])
def student_submit():
    # Student details
    reg_no = request.form["regno"]
    name = request.form["name"]
    email = request.form["email"]
    phone = request.form["phone"]
    department = request.form["department"]
    year = request.form["year"]

    # Internship details
    title = request.form["domain"]
    company = request.form["company"]
    duration = request.form["duration"]
    start_date = request.form["start_date"]
    end_date = request.form["end_date"]
    address = request.form["address"]

    # File upload
    file = request.files["offer_letter"]
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    # Save into Excel
    new_data = pd.DataFrame([[
        reg_no, name, email, phone, department, year,
        title, company, address, duration, start_date, end_date,
        file.filename, "pending"
    ]], columns=[
        "Register No", "Name", "Email", "Phone", "Department", "Year",
        "Internship Title", "Company", "Location", "Duration",
        "Start Date", "End Date", "Offer Letter", "Status"
    ])

    if os.path.exists(EXCEL_FILE):
        old_data = pd.read_excel(EXCEL_FILE)
        new_data = pd.concat([old_data, new_data], ignore_index=True)

    new_data.to_excel(app.config["EXCEL_FILE"], index=False)

    return render_template("thankyou.html", name=name)


@app.route("/coordinator_login", methods=["GET", "POST"])
def coordinator_login():
    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")

        # Example: static credentials (replace with DB later)
        if username == "admin" and password == "1234":
            session["coordinator"] = username
            return redirect(url_for("coordinator_dashboard"))
        else:
            return render_template("coordinator_login.html", error="Invalid credentials")

    return render_template("coordinator_login.html")


@app.route("/coordinator")
def coordinator_dashboard():
    import pandas as pd
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        data = df.to_dict(orient="records")  # convert to list of dicts
    else:
        data = []

    return render_template("coordinator.html", data=data)




@app.route('/update_status/<int:row_index>/<status>')
def update_status(row_index, status):
    df = pd.read_excel("internship_data.xlsx")
    if 0 <= row_index < len(df):
        df.at[row_index, 'Status'] = status
        df.to_excel("internship_data.xlsx", index=False)
    return redirect(url_for('coordinator'))


@app.route("/uploads/<filename>")
def uploaded_file(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)

@app.route("/approve/<int:index>", methods=["POST"])
def approve_submission(index):
    df = pd.read_excel(EXCEL_FILE)
    df.loc[index, "Status"] = "Approved"
    df.to_excel(EXCEL_FILE, index=False)
    return redirect(url_for("coordinator_dashboard"))

@app.route("/reject/<int:index>", methods=["POST"])
def reject_submission(index):
    df = pd.read_excel(EXCEL_FILE)
    df.loc[index, "Status"] = "Rejected"
    df.to_excel(EXCEL_FILE, index=False)
    return redirect(url_for("coordinator_dashboard"))


@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)



if __name__ == '__main__':
    app.run(debug=True)

