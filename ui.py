import sys
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                              QHBoxLayout, QPushButton, QComboBox, QLabel,
                              QMessageBox, QGridLayout, QScrollArea)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor, QPalette
from scheduling_logic import (EMPLOYEES, FREELANCERS, SHIFT_COLORS, 
                              load_data, save_data, init_availability, 
                              generate_schedule, import_from_excel, 
                              export_availability_to_excel, clear_availability)
from datetime import datetime

class AvailabilityEditor(QMainWindow):
    def __init__(self, start_date=datetime(2025, 3, 17)):
        super().__init__()
        self.start_date = start_date
        self.employees = EMPLOYEES
        self.employee_names = FREELANCERS
        self.current_employee_name = self.employee_names[0] if self.employee_names else ""
        
        loaded_data = load_data()
        self.availability = loaded_data if loaded_data else init_availability(start_date, self.employees)
        
        self.init_ui()
        self.update_calendar()

    def init_ui(self):
        self.setWindowTitle("Employee Availability Editor")
        self.setGeometry(100, 100, 1000, 600)
        
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        control_layout = QHBoxLayout()
        self.employee_combo = QComboBox()
        self.employee_combo.addItems(self.employee_names)
        self.employee_combo.currentTextChanged.connect(self.employee_changed)
        control_layout.addWidget(QLabel("Select Employee:"))
        control_layout.addWidget(self.employee_combo)
        
        scroll = QScrollArea()
        self.calendar_widget = QWidget()
        self.calendar_layout = QGridLayout(self.calendar_widget)
        scroll.setWidget(self.calendar_widget)
        scroll.setWidgetResizable(True)
        
        button_layout = QHBoxLayout()
        save_btn = QPushButton("保存時間")
        save_btn.clicked.connect(self.save_data)
        generate_btn = QPushButton("導出至Excel")
        generate_btn.clicked.connect(self.generate_schedule)
        import_btn = QPushButton("從Excel導入")
        import_btn.clicked.connect(lambda: self.import_from_excel("availability_export.xlsx"))
        export_btn = QPushButton("導出時間表至Excel")
        export_btn.clicked.connect(self.export_availability_to_excel)
        clear_btn = QPushButton("Clear Availability")
        clear_btn.clicked.connect(self.clear_availability)
        
        button_layout.addWidget(save_btn)
        button_layout.addWidget(generate_btn)
        button_layout.addWidget(import_btn)
        button_layout.addWidget(export_btn)
        button_layout.addWidget(clear_btn)
        
        layout.addLayout(control_layout)
        layout.addWidget(scroll)
        layout.addLayout(button_layout)

    def create_day_widget(self, date_str):
        day_widget = QWidget()
        day_layout = QVBoxLayout(day_widget)
        
        date = datetime.strptime(date_str, "%Y-%m-%d")
        day_label = QLabel(date.strftime("%a\n%d %b"))
        day_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        day_layout.addWidget(day_label)
        
        current_employee = next((e for e in self.employees if e.get_display_name() == self.current_employee_name), None)
        
        if current_employee:
            available_shifts = current_employee.get_available_shifts()
            
            for shift in available_shifts:
                if shift in SHIFT_COLORS:
                    color = QColor(*SHIFT_COLORS[shift])
                    btn = QPushButton(shift)
                    btn.setCheckable(True)
                    btn.setStyleSheet(f"""
                        QPushButton {{ 
                            background-color: {color.name()}; 
                            border: 1px solid gray;
                            color: white;  
                        }}
                        QPushButton:checked {{
                            border: 2px solid white;
                            color: black;
                        }}
                    """)
                    btn.clicked.connect(lambda _, d=date_str, s=shift: self.toggle_shift(d, s))
                    day_layout.addWidget(btn)
        
        return day_widget

    def update_calendar(self):
        for i in reversed(range(self.calendar_layout.count())): 
            self.calendar_layout.itemAt(i).widget().setParent(None)
        
        sorted_dates = sorted(self.availability.keys())
        for col, date_str in enumerate(sorted_dates):
            day_widget = self.create_day_widget(date_str)
            self.calendar_layout.addWidget(day_widget, 0, col)
            
            if self.current_employee_name in self.availability[date_str]:
                shifts = self.availability[date_str][self.current_employee_name]
                for i in range(day_widget.layout().count()):
                    widget = day_widget.layout().itemAt(i).widget()
                    if isinstance(widget, QPushButton):
                        widget.setChecked(widget.text() in shifts)

    def toggle_shift(self, date_str, shift):
        if self.current_employee_name in self.availability[date_str]:
            current_shifts = self.availability[date_str][self.current_employee_name]
            if shift in current_shifts:
                current_shifts.remove(shift)
            else:
                current_shifts.append(shift)
            
            save_data(self.availability)
            self.update_calendar()

    def employee_changed(self, name):
        self.current_employee_name = name
        self.update_calendar()

    def save_data(self):
        save_data(self.availability)
        QMessageBox.information(self, "Saved", "時間已成功保存!")

    def generate_schedule(self):
        try:
            result = generate_schedule(self.availability, self.start_date)
            QMessageBox.information(self, "Success", result)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error: {str(e)}")

    def import_from_excel(self, file_path):
        try:
            result = import_from_excel(file_path)
            self.availability = load_data()  # Reload the data after import
            self.update_calendar()  # Update the UI with the new data
            QMessageBox.information(self, "Success", result)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to import data: {str(e)}")

    def export_availability_to_excel(self):
        try:
            result = export_availability_to_excel(self.availability)
            QMessageBox.information(self, "Success", result)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export availability: {str(e)}")

    def clear_availability(self):
        reply = QMessageBox.question(
            self,
            "Confirm Clear",
            "Are you sure you want to clear all stored availability?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.availability = clear_availability(self.start_date, self.employees)
            save_data(self.availability)
            self.update_calendar()
            QMessageBox.information(self, "Cleared", "All availability data has been cleared!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AvailabilityEditor()
    window.show()
    sys.exit(app.exec())
