import sys
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                              QHBoxLayout, QPushButton, QComboBox, QLabel,
                              QMessageBox, QGridLayout, QScrollArea, QDialog, QLineEdit, QMenu)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor, QPalette
from scheduling_logic import (EMPLOYEES, Freelancer, SHIFT_COLORS, 
                              load_data, save_data, init_availability, 
                               generate_schedule, import_from_excel, 
                               edit_employee, load_employees, ROLE_RULES, add_employee, delete_employee,
                              export_availability_to_excel, clear_availability)
from datetime import datetime

class AvailabilityEditor(QMainWindow):
    def __init__(self, start_date=datetime(2025, 3, 17)):
        super().__init__()
        self.start_date = start_date
        self.employees = load_employees()  # Load employees from JSON
        self.employee_names = [emp.get_display_name() for emp in self.employees if isinstance(emp, Freelancer)]
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

        # Control layout for dropdowns and buttons
        control_layout = QHBoxLayout()
        self.employee_combo = QComboBox()
        self.employee_combo.addItems(self.employee_names)
        self.employee_combo.currentTextChanged.connect(self.select_employee)

        # Role selection
        self.role_combo = QComboBox()
        self.role_combo.addItems(["All"] + list(ROLE_RULES.keys()))
        self.role_combo.currentTextChanged.connect(self.role_changed)
        control_layout.addWidget(QLabel("Select Role:"))
        control_layout.addWidget(self.role_combo)

        # Add employee list
        self.employee_list = QWidget()
        self.employee_list_layout = QVBoxLayout(self.employee_list)
        control_layout.addWidget(QLabel("Select Employee:"))
        control_layout.addWidget(self.employee_list)

        layout.addLayout(control_layout)

        # Create a container for the calendar and selected employee label
        calendar_container = QVBoxLayout()

        # Move "Currently Selected" label here
        self.selected_employee_label = QLabel(f"Currently Selected: {self.current_employee_name}")
        self.selected_employee_label.setAlignment(Qt.AlignmentFlag.AlignLeft)  # Align to top-left
        calendar_container.addWidget(self.selected_employee_label)

        # Scroll area for the calendar
        scroll = QScrollArea()
        self.calendar_widget = QWidget()
        self.calendar_layout = QGridLayout(self.calendar_widget)
        scroll.setWidget(self.calendar_widget)
        scroll.setWidgetResizable(True)
        
        # Add scroll area to calendar container
        calendar_container.addWidget(scroll)

        # Add calendar container to main layout
        layout.addLayout(calendar_container)

        # Button layout for actions
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

        validate_btn = QPushButton("Validate Schedule")
        validate_btn.clicked.connect(self.validate_schedule)

        add_btn = QPushButton("Add Employee")
        add_btn.clicked.connect(self.add_new_employee)

        button_layout.addWidget(save_btn)
        button_layout.addWidget(generate_btn)
        button_layout.addWidget(import_btn)
        button_layout.addWidget(export_btn)
        button_layout.addWidget(clear_btn)
        button_layout.addWidget(validate_btn)
        button_layout.addWidget(add_btn)

        layout.addLayout(button_layout)

    def add_new_employee(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add New Employee")
            
        layout = QVBoxLayout(dialog)
            
        name_input = QLineEdit()
        role_input = QComboBox()
        role_input.addItems(list(ROLE_RULES.keys()))
            
        layout.addWidget(QLabel("Name:"))
        layout.addWidget(name_input)
        layout.addWidget(QLabel("Role:"))
        layout.addWidget(role_input)
            
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(lambda: self.save_new_employee(dialog, name_input.text(), role_input.currentText()))
            
        layout.addWidget(save_btn)
            
        dialog.exec_()

    def save_new_employee(self, dialog, name, role):
        add_employee(name, role)
        self.employees = load_employees()
        self.update_employee_list(self.role_combo.currentText())
        dialog.accept()


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

    def role_changed(self, role):
        self.update_employee_list(role)

    def update_employee_list(self, role):
        # Clear existing items in the list
        for i in reversed(range(self.employee_list_layout.count())):
            self.employee_list_layout.itemAt(i).widget().setParent(None)

        # Filter employees by role
        filtered_employees = [emp.get_display_name() for emp in self.employees if role == "All" or emp.employee_type == role]

        for name in filtered_employees:
            btn = QPushButton(name)
            btn.clicked.connect(lambda _, n=name: self.select_employee(n))

            # Highlight the currently selected employee
            if name == self.current_employee_name:
                btn.setStyleSheet("background-color: lightblue; font-weight: bold;")
            else:
                btn.setStyleSheet("")  # Reset style for non-selected employees

            btn.setContextMenuPolicy(Qt.CustomContextMenu)
            btn.customContextMenuRequested.connect(
                lambda pos, btn=btn, n=name: self.show_context_menu(pos, btn, n)
            )
            self.employee_list_layout.addWidget(btn)





    def select_employee(self, name):
        self.current_employee_name = name
        self.selected_employee_label.setText(f"Currently Selected: {self.current_employee_name}")  # Update label
        self.update_calendar()  # Refresh calendar view
        self.update_employee_list(self.role_combo.currentText())  # Refresh employee list to highlight selection



    def show_context_menu(self, pos, button, name):  # Modified signature
        menu = QMenu()
        edit_action = menu.addAction("Edit")
        delete_action = menu.addAction("Delete")
        
        # Use the passed button reference instead of sender()
        action = menu.exec_(button.mapToGlobal(pos))  # Critical fix
        
        if action == edit_action:
            self.edit_employee(name)
        elif action == delete_action:
            self.confirm_delete(name)

    def edit_employee(self, name):
        old_employee = next((emp for emp in self.employees if emp.get_display_name() == name), None)
        
        if old_employee:
            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Employee")
            
            layout = QVBoxLayout(dialog)
            
            name_input = QLineEdit(old_employee.name)
            role_input = QComboBox()
            role_input.addItems(list(ROLE_RULES.keys()))
            role_input.setCurrentText(old_employee.employee_type)
            
            layout.addWidget(QLabel("Name:"))
            layout.addWidget(name_input)
            layout.addWidget(QLabel("Role:"))
            layout.addWidget(role_input)
            
            save_btn = QPushButton("Save")
            save_btn.clicked.connect(lambda: self.save_edited_employee(dialog, old_employee.name, name_input.text(), role_input.currentText()))
            
            layout.addWidget(save_btn)
            
            dialog.exec_()

    def save_edited_employee(self, dialog, old_name, new_name, new_role):
        edit_employee(old_name, new_name, new_role)
        self.employees = load_employees()
        self.update_employee_list(self.role_combo.currentText())
        dialog.accept()


    def save_data(self):
        save_data(self.availability)
        QMessageBox.information(self, "Saved", "時間已成功保存!")
    
    def validate_schedule(self):
        warnings = generate_schedule(self.availability, self.start_date, export_to_excel=False)
        if warnings:
            warning_message = "\n".join(warnings)
            QMessageBox.warning(self, "Scheduling Warnings", warning_message)
        else:
            QMessageBox.information(self, "Validation", "No issues found with the current schedule.")





    def generate_schedule(self):
        try:
            warnings = generate_schedule(self.availability, self.start_date)
            if warnings:
                warning_message = "\n".join(warnings)
                QMessageBox.warning(self, "Scheduling Warnings", warning_message)
            QMessageBox.information(self, "Success", "Excel schedule has been successfully generated!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate schedule: {str(e)}")



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
    def confirm_delete(self, name):
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete {name}?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            delete_employee(name)
            self.employees = load_employees()
            self.update_employee_list(self.role_combo.currentText())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AvailabilityEditor()
    window.show()
    sys.exit(app.exec())
