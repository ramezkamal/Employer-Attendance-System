import sys
import os
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QLabel, QLineEdit, 
                             QCheckBox, QVBoxLayout, QWidget, QDialog, QMessageBox, 
                             QTableWidget, QTableWidgetItem, QHBoxLayout, QComboBox, QFormLayout)
from PyQt5.QtGui import QIcon, QFont, QColor, QPixmap
from PyQt5.QtCore import Qt

# الأنماط البصرية
STYLE_SHEET = """
    QWidget {
        font-family: 'Segoe UI';
        font-size: 20px;
        background-color:rgb(45, 157, 255);
    }
    
    QPushButton {
        background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #FF6F61, stop:1 #FF9E80);
        color: white;
        padding: 25px 50px;
        border-radius: 20px;
        font-size: 22px;
        min-width: 350px;
        min-height: 80px;
        margin: 25px;
        text-align: center;
        box-shadow: 5px 5px 10px rgba(0,0,0,0.2);
    }
    
    QPushButton:hover {
        background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #FF9E80, stop:1 #FF6F61);
    }
    
    QLineEdit, QComboBox {
        padding: 20px;
        border: 4px solid #FF6F61;
        border-radius: 15px;
        font-size: 20px;
        min-width: 500px;
        min-height: 70px;
        background-color: #FFF3E0;
        color: #27374D;
    }
    
    QLabel {
        font-size: 24px;
        color: #27374D;
        margin: 5px;
    }
    
    QCheckBox {
        spacing: 20px;
        font-size: 20px;
    }
    
    QTableWidget {
        background-color: white;
        border: 3px solid #FF6F61;
        border-radius: 15px;
        margin: 25px;
        font-size: 20px;
    }
    
    QComboBox {
        padding: 20px;
        min-height: 70px;
    }
"""

# تعريف كلاس MainWindow أولاً
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("نظام إدارة الحضور")
        self.setGeometry(100, 100, 1200, 900)
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()
        
        # العناوين الرئيسية
        title_layout = QVBoxLayout()
        
        bismillah_label = QLabel("✨ *بسم الثالوث القدوس* ✨")
        bismillah_label.setFont(QFont('Segoe UI', 80, QFont.Bold))
        bismillah_label.setAlignment(Qt.AlignCenter)
        bismillah_label.setStyleSheet("color:rgb(255, 208, 0); margin: 0px;")
        
        database_label = QLabel("🌟 قاعدة بيانات مركز 🌟")
        database_label.setFont(QFont('Segoe UI', 60,QFont.Bold))
        database_label.setAlignment(Qt.AlignCenter)
        database_label.setStyleSheet("color:rgb(255, 204, 0); margin: 0px;")
        
        clinic_name = QLabel("🏥 افابيجول الطبي 🏥")
        clinic_name.setFont(QFont('Segoe UI',50,QFont.Bold))
        clinic_name.setAlignment(Qt.AlignCenter)
        clinic_name.setStyleSheet("color:rgb(255, 196, 0); margin: 0px;")
        
        title_layout.addWidget(bismillah_label)
        title_layout.addWidget(database_label)
        title_layout.addWidget(clinic_name)
        title_layout.setSpacing(0)
        
        stars_label = QLabel("⭐️ ⭐️ ⭐️")
        stars_label.setFont(QFont('Segoe UI', 20))
        stars_label.setAlignment(Qt.AlignCenter)
        stars_label.setStyleSheet("color: #FFD700; margin: 5px;")
        
        # الأزرار الرئيسية
        self.admin_btn = QPushButton("🔒 تسجيل كأدمن ✨")
        self.admin_btn.setIcon(QIcon('icons/download.png'))
        self.admin_btn.setStyleSheet("""
            QPushButton {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #FFAFBD, stop:1 #ffc3a0);
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #FFC3A0, stop:1 #FFAFBD);
            }
        """)
        
        self.attendance_btn = QPushButton("📝 بدء تسجيل الحضور/الانصراف 🕘")
        self.attendance_btn.setIcon(QIcon('icons/attendance.png'))
        self.attendance_btn.setStyleSheet("""
            QPushButton {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #8EC5FC, stop:1 #E0C3FC);
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #E0C3FC, stop:1 #8EC5FC);
            }
        """)
        
        layout.addLayout(title_layout)
        layout.addWidget(stars_label)
        layout.addWidget(self.admin_btn)
        layout.addWidget(self.attendance_btn)
        
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        
        # ربط الأحداث
        self.admin_btn.clicked.connect(self.show_admin_login)
        self.attendance_btn.clicked.connect(self.show_attendance)
        
    def show_admin_login(self):
        dialog = AdminLogin()
        if dialog.exec_() == QDialog.Accepted:
            self.employee_window = EmployeeManagement()
            self.employee_window.show()
            
    def show_attendance(self):
        self.attendance_window = AttendanceWindow()
        self.attendance_window.show()

# كلاس تسجيل الدخول للأدمن
class AdminLogin(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("دخول الأدمن")
        self.setGeometry(300, 200, 600, 500)
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()
        
        self.username = QLineEdit()
        self.username.setPlaceholderText("اسم المستخدم 📝")
        self.username.setFont(QFont('Arial', 20))
        
        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.Password)
        self.password.setPlaceholderText("كلمة المرور 🔒")
        self.password.setFont(QFont('Arial', 20))
        
        login_btn = QPushButton("تسجيل الدخول ✨")
        login_btn.setStyleSheet("background-color: #3498DB;")
        login_btn.setFont(QFont('Arial', 22))
        
        layout.addWidget(QLabel("تسجيل دخول الأدمن"))
        layout.addWidget(self.username)
        layout.addWidget(self.password)
        layout.addWidget(login_btn)
        
        self.setLayout(layout)
        login_btn.clicked.connect(self.verify)
        
    def verify(self):
        if self.username.text() == 'admin' and self.password.text() == 'admiiin':
            self.accept()
        else:
            QMessageBox.warning(self, "خطأ", "بيانات الدخول غير صحيحة")

# كلاس إدارة الموظفين
class EmployeeManagement(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("إدارة الموظفين")
        self.setGeometry(200, 100, 1200, 800)
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()
        
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(['الاسم', 'يونيفورم', 'رقم الهاتف', 'تاريخ الإضافة'])
        self.table.horizontalHeader().setStyleSheet("font-weight: bold;")
        self.load_employees()
        
        add_btn = QPushButton("إضافة موظف جديد 🆕")
        add_btn.setStyleSheet("background-color: #2ECC71;")
        add_btn.setIcon(QIcon('icons/add.png'))
        add_btn.setFont(QFont('Arial', 20))
        
        layout.addWidget(self.table)
        layout.addWidget(add_btn)
        
        self.setLayout(layout)
        add_btn.clicked.connect(self.show_add_dialog)
        
    def load_employees(self):
        if os.path.exists('Database.xlsx'):
            df = pd.read_excel('Database.xlsx')
            self.table.setRowCount(df.shape[0])
            for i, row in df.iterrows():
                for j in range(self.table.columnCount()):
                    self.table.setItem(i, j, QTableWidgetItem(str(row.iloc[j])))
        else:
            QMessageBox.information(self, "ملاحظة", "لا يوجد موظفون مسجلون")

    def show_add_dialog(self):
        dialog = AddEmployeeDialog()
        if dialog.exec_():
            self.load_employees()

# كلاس إضافة موظف جديد
class AddEmployeeDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("إضافة موظف جديد")
        self.setGeometry(300, 200, 800, 600)
        self.initUI()
        
    def initUI(self):
        layout = QFormLayout()
        
        self.name = QLineEdit()
        self.name.setPlaceholderText("اسم الموظف 📝")
        self.name.setFont(QFont('Arial', 20))
        
        self.uniform = QComboBox()
        self.uniform.addItems(['نعم ✅', 'لا ❌'])
        self.uniform.setFont(QFont('Arial', 20))
        
        self.phone = QLineEdit()
        self.phone.setPlaceholderText("رقم الهاتف 📞")
        self.phone.setFont(QFont('Arial', 20))
        
        save_btn = QPushButton("حفظ الموظف 💾")
        save_btn.setStyleSheet("background-color: #3498DB;")
        save_btn.setFont(QFont('Arial', 22))
        
        layout.addRow("الاسم:", self.name)
        layout.addRow("يونيفورم:", self.uniform)
        layout.addRow("الهاتف:", self.phone)
        layout.addRow(save_btn)
        
        self.setLayout(layout)
        save_btn.clicked.connect(self.save_employee)
        
    def save_employee(self):
        name = self.name.text().strip()
        uniform = self.uniform.currentText()
        phone = self.phone.text().strip()
        date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if not name or not phone:
            QMessageBox.warning(self, "خطأ", "يرجى إدخال جميع البيانات")
            return
            
        try:
            new_data = pd.DataFrame([[name, uniform, phone, date]], 
                                   columns=['الاسم', 'يونيفورم', 'رقم الهاتف', 'تاريخ الإضافة'])
            
            if os.path.exists('Database.xlsx'):
                df = pd.read_excel('Database.xlsx')
                df = pd.concat([df, new_data], ignore_index=True)
            else:
                df = new_data
                
            df.to_excel('Database.xlsx', index=False)
            QMessageBox.information(self, "نجاح", "تم إضافة الموظف بنجاح")
            self.accept()
            
        except Exception as e:
            QMessageBox.critical(self, "خطأ", str(e))

# كلاس تسجيل الحضور والانصراف
class AttendanceWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("تسجيل الحضور والانصراف")
        self.setGeometry(200, 100, 1200, 800)
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()
        
        self.employee_combo = QComboBox()
        self.employee_combo.setFont(QFont('Arial', 20))
        self.load_employees()
        
        self.status_combo = QComboBox()
        self.status_combo.addItems(['حضور ✅', 'انصراف ❌'])
        self.status_combo.setFont(QFont('Arial', 20))
        
        self.uniform_combo = QComboBox()
        self.uniform_combo.addItems(['يرتدي 🎽', 'لا يرتدي ❌'])
        self.uniform_combo.setFont(QFont('Arial', 20))
        
        register_btn = QPushButton("تسجيل 📝")
        register_btn.setStyleSheet("background-color: #E74C3C;")
        register_btn.setFont(QFont('Arial', 22))
        register_btn.setIcon(QIcon('icons/check.png'))
        
        form_layout = QFormLayout()
        form_layout.addRow("اسم الموظف:", self.employee_combo)
        form_layout.addRow("الحالة:", self.status_combo)
        form_layout.addRow("الزي الرسمي:", self.uniform_combo)
        
        layout.addLayout(form_layout)
        layout.addWidget(register_btn)
        
        self.setLayout(layout)
        register_btn.clicked.connect(self.record_attendance)
        
    def load_employees(self):
        if os.path.exists('Database.xlsx'):
            df = pd.read_excel('Database.xlsx')
            self.employee_combo.addItems(df['الاسم'].tolist())
        else:
            self.employee_combo.addItem("لا يوجد موظفون مسجلون")
        
    def record_attendance(self):
        employee = self.employee_combo.currentText()
        status = self.status_combo.currentText()
        uniform = self.uniform_combo.currentText()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if employee == "لا يوجد موظفون مسجلون":
            QMessageBox.warning(self, "خطأ", "لا يوجد موظفون مسجلون")
            return
            
        try:
            current_month = datetime.now().strftime("%Y-%m")
            file_name = f"{current_month}.xlsx"
            
            if os.path.exists(file_name):
                df = pd.read_excel(file_name)
            else:
                df = pd.DataFrame(columns=['الاسم', 'الحضور', 'الانصراف'])
            
            if employee not in df['الاسم'].values:
                new_row = {'الاسم': employee, 'الحضور': None, 'الانصراف': None}
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            
            column = 'الحضور' if '✅' in status else 'الانصراف'
            df.loc[df['الاسم'] == employee, column] = f"{uniform} ({timestamp})"
            
            df.to_excel(file_name, index=False)
            QMessageBox.information(self, "نجاح", "تم التسجيل بنجاح")
            
        except Exception as e:
            QMessageBox.critical(self, "خطأ", str(e))

# تشغيل التطبيق
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLE_SHEET)
    
    if not os.path.exists("icons"):
        os.makedirs("icons")
    
    def create_icon(name, color):
        pixmap = QPixmap(64, 64)
        pixmap.fill(QColor(color))
        icon = QIcon(pixmap)
        icon.pixmap(64, 64).save(f"icons/{name}.png")
    
    create_icon("admin", "#1ABC9C")
    create_icon("attendance", "#F39C12")
    create_icon("add", "#2ECC71")
    create_icon("check", "#E74C3C")
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())