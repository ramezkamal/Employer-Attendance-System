from flask import Flask, render_template, request, redirect, url_for, flash, session
import pandas as pd
from datetime import datetime
import os
from pathlib import Path

app = Flask(__name__)
app.secret_key = 'your_secret_key'
app.config['UPLOAD_FOLDER'] = 'static'

# تحديد المسار الأساسي
BASE_DIR = Path(__file__).parent

# تأمين الدخول
def login_required(f):
    def wrap(*args, **kwargs):
        if 'logged_in' in session:
            return f(*args, **kwargs)
        else:
            return redirect(url_for('admin_login'))
    return wrap

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/admin_login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if username == 'admin' and password == '123':
            session['logged_in'] = True
            return redirect(url_for('admin_dashboard'))
        else:
            flash('بيانات الدخول غير صحيحة', 'error')
    return render_template('admin_login.html')

@app.route('/admin_dashboard', methods=['GET', 'POST'])
@login_required
def admin_dashboard():
    if request.method == 'POST':
        name = request.form['name']
        uniform = request.form['uniform']
        phone = request.form['phone']
        date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        new_data = pd.DataFrame([[name, uniform, phone, date]], 
                                columns=['الاسم', 'يونيفورم', 'رقم الهاتف', 'تاريخ الإضافة'])
        
        db_path = BASE_DIR / 'Database.xlsx'
        if db_path.exists():
            df = pd.read_excel(db_path)
            df = pd.concat([df, new_data], ignore_index=True)
        else:
            df = new_data
        df.to_excel(db_path, index=False)
        flash('تم إضافة الموظف بنجاح', 'success')
    
    employees = []
    db_path = BASE_DIR / 'Database.xlsx'
    if db_path.exists():
        df = pd.read_excel(db_path)
        employees = df.to_dict('records')
    return render_template('admin_dashboard.html', employees=employees)

@app.route('/attendance', methods=['GET', 'POST'])
def attendance():
    if request.method == 'POST':
        employee = request.form['employee']
        status = request.form['status']
        uniform = request.form['uniform']
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        current_month = datetime.now().strftime("%Y-%m")
        file_path = BASE_DIR / f"{current_month}.xlsx"
        
        new_data = pd.DataFrame([[employee, status, uniform, timestamp]],
                                columns=['الاسم', 'الحالة', 'الزي', 'التاريخ'])
        
        if file_path.exists():
            df = pd.read_excel(file_path)
            df = pd.concat([df, new_data], ignore_index=True)
        else:
            df = new_data
        df.to_excel(file_path, index=False)
        flash('تم التسجيل بنجاح', 'success')
    
    employees = []
    db_path = BASE_DIR / 'Database.xlsx'
    if db_path.exists():
        df = pd.read_excel(db_path)
        employees = df['الاسم'].tolist()
    return render_template('attendance.html', employees=employees)

if __name__ == '__main__':
    app.run(debug=True)