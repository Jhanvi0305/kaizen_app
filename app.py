from flask import Flask, render_template, request, redirect, send_file
import datetime
import sqlite3
import os
from openpyxl import Workbook

app = Flask(__name__)

UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# ✅ Ensure upload folder exists (FIXED POSITION)
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# DB INIT
def init_db():
    conn = sqlite3.connect('/tmp/kaizen.db')
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS kaizens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kaizen_id TEXT,
            emp1_name TEXT,
            emp2_name TEXT,
            emp3_name TEXT,
            department TEXT,
            area TEXT,
            title TEXT,
            before_text TEXT,
            after_text TEXT,
            category TEXT,
            monthly_saving REAL,
            yearly_saving REAL,
            one_time_saving REAL,
            before_file TEXT,
            after_file TEXT,
            status TEXT,
            sustenance_status TEXT,
            last_updated TEXT,
            created_at TEXT
        )
    ''')

    conn.commit()
    conn.close()

init_db()

# HOME
@app.route('/')
def home():
    return render_template('index.html')

# SUBMIT
@app.route('/submit', methods=['POST'])
def submit():

    emp1 = request.form.get('emp1_name')
    emp2 = request.form.get('emp2_name')
    emp3 = request.form.get('emp3_name')

    department = request.form.get('department')
    area = request.form.get('area')
    title = request.form.get('title')
    before = request.form.get('before')
    after = request.form.get('after')
    category = request.form.get('category')

    monthly = request.form.get('monthly_saving')
    yearly = request.form.get('yearly_saving')
    one_time = request.form.get('one_time_saving')

    before_file = request.files.get('before_file')
    after_file = request.files.get('after_file')

    before_filename = ""
    after_filename = ""

    if before_file and before_file.filename:
        before_filename = str(datetime.datetime.now().timestamp()) + "_" + before_file.filename
        before_file.save(os.path.join(app.config['UPLOAD_FOLDER'], before_filename))

    if after_file and after_file.filename:
        after_filename = str(datetime.datetime.now().timestamp()) + "_" + after_file.filename
        after_file.save(os.path.join(app.config['UPLOAD_FOLDER'], after_filename))

    kaizen_id = "KZ-" + datetime.datetime.now().strftime("%Y%m%d%H%M%S")

    conn = sqlite3.connect('/tmp/kaizen.db')
    cursor = conn.cursor()

    cursor.execute('''
        INSERT INTO kaizens 
        (kaizen_id, emp1_name, emp2_name, emp3_name, department, area, title, before_text, after_text, category, monthly_saving, yearly_saving, one_time_saving, before_file, after_file, status, sustenance_status, last_updated, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        kaizen_id, emp1, emp2, emp3,
        department, area, title,
        before, after,
        category, monthly, yearly, one_time,
        before_filename, after_filename,
        "Pending",
        "Red",
        datetime.datetime.now(),
        datetime.datetime.now()
    ))

    conn.commit()
    conn.close()

    return "✅ Kaizen Submitted Successfully!"

# DASHBOARD
@app.route('/dashboard')
def dashboard():

    conn = sqlite3.connect('/tmp/kaizen.db')
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM kaizens")
    data = cursor.fetchall()

    total_kaizens = len(data)
    total_savings = 0
    category_data = {}
    dept_data = {}

    for row in data:
        try:
            total_savings += float(row[11]) if row[11] else 0
        except:
            pass

        cat = row[10]
        if cat:
            category_data[cat] = category_data.get(cat, 0) + 1

        dept = row[5]
        if dept:
            dept_data[dept] = dept_data.get(dept, 0) + 1

    conn.close()

    return render_template(
        'dashboard.html',
        data=data,
        total_kaizens=total_kaizens,
        total_savings=total_savings,
        category_data=category_data,
        dept_data=dept_data
    )

# APPROVE
@app.route('/approve/<int:id>')
def approve(id):
    conn = sqlite3.connect('/tmp/kaizen.db')
    cursor = conn.cursor()
    cursor.execute("UPDATE kaizens SET status='Approved' WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return redirect('/dashboard')

# REJECT
@app.route('/reject/<int:id>')
def reject(id):
    conn = sqlite3.connect('/tmp/kaizen.db')
    cursor = conn.cursor()
    cursor.execute("UPDATE kaizens SET status='Rejected' WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return redirect('/dashboard')

# SUSTENANCE
@app.route('/update_sustenance/<int:id>', methods=['POST'])
def update_sustenance(id):

    file = request.files.get('proof_file')

    if file and file.filename:
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], file.filename))

    conn = sqlite3.connect('/tmp/kaizen.db')
    cursor = conn.cursor()

    cursor.execute('''
        UPDATE kaizens 
        SET sustenance_status='Green', last_updated=?
        WHERE id=?
    ''', (datetime.datetime.now(), id))

    conn.commit()
    conn.close()

    return redirect('/dashboard')

# EXPORT EXCEL
@app.route('/export')
def export():

    conn = sqlite3.connect('/tmp/kaizen.db')
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM kaizens")
    data = cursor.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Kaizen Report"

    headers = [
        "Kaizen ID", "Emp1", "Emp2", "Emp3",
        "Department", "Category", "Monthly Saving",
        "Status", "Sustenance"
    ]
    ws.append(headers)

    for row in data:
        ws.append([
            row[1], row[2], row[3], row[4],
            row[5], row[10], row[11],
            row[16], row[17]
        ])

    file_path = "kaizen_report.xlsx"
    wb.save(file_path)

    return send_file(file_path, as_attachment=True)

# RUN APP
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
