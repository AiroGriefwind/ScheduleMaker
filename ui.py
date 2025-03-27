import sys
import pandas as pd # Import for SchedulePreviewDialog
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                              QHBoxLayout, QPushButton, QComboBox, QLabel,
                              QMessageBox, QGridLayout, QScrollArea, QDialog, QLineEdit, QMenu)
from PySide6.QtWidgets import QTableWidget, QTableWidgetItem, QHeaderView, QDialogButtonBox # Import for SchedulePreviewDialog
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor, QPalette
from scheduling_logic import (EMPLOYEES, Freelancer,  
                              load_data, save_data, init_availability, 
                               generate_schedule, import_from_excel, 
                               edit_employee, load_employees, ROLE_RULES, add_employee, delete_employee,sync_availability,
                              export_availability_to_excel, clear_availability)
from datetime import datetime

class AvailabilityEditor(QMainWindow):
    def __init__(self, start_date=datetime(2025, 3, 17)):
        super().__init__()
        self.start_date = start_date
        self.employees = load_employees()  # Load employees from JSON
        self.employee_names = [emp.name for emp in self.employees if isinstance(emp, Freelancer)]
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

        button_layout = QHBoxLayout()# Add button layout for 從google表單導入

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

        # Create a scroll area for employee list
        employee_scroll = QScrollArea()
        employee_scroll.setWidgetResizable(True)
        employee_scroll.setFixedHeight(200)  # Set fixed height here
        employee_scroll.setMinimumWidth(50)  # Minimum width
        employee_scroll.setMaximumWidth(150)  # Maximum width

        
        # Create employee list widget and layout
        self.employee_list = QWidget()
        self.employee_list_layout = QVBoxLayout(self.employee_list)
        self.employee_list_layout.setContentsMargins(5, 5, 5, 5)
        
        # Add employee list to scroll area
        employee_scroll.setWidget(self.employee_list)
        
        # Add label and scroll area to control layout
        control_layout.addWidget(QLabel("Select Employee:"))
        control_layout.addWidget(employee_scroll)

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

        import_form_btn = QPushButton("從Google表單導入")
        import_form_btn.clicked.connect(self.import_from_google_form)

        button_layout.addWidget(save_btn)
        button_layout.addWidget(generate_btn)
        button_layout.addWidget(import_btn)
        button_layout.addWidget(import_form_btn)  # Add the button for Google Form import
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

    # Updated ui.py
    def save_new_employee(self, dialog, name, role):
        add_employee(name, role)
        self.employees = load_employees()
        self.availability = load_data()  # Critical reload
        self.update_employee_list(self.role_combo.currentText())
        self.update_calendar()  # Refresh UI components
        dialog.accept()



    def create_day_widget(self, date_str):
        day_widget = QWidget()
        day_layout = QVBoxLayout(day_widget)
        
        date = datetime.strptime(date_str, "%Y-%m-%d")
        day_label = QLabel(date.strftime("%a\n%d %b"))
        day_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        day_layout.addWidget(day_label)
        
        current_employee = next((e for e in self.employees if e.name == self.current_employee_name), None)
        
        if current_employee:
            role = current_employee.employee_type
            available_shifts = []
            
            if role in ROLE_RULES:
                rule = ROLE_RULES[role]
                if rule["rule_type"] == "shift_based":
                    day_type = "weekend" if date.weekday() >= 5 else "weekday"
                    available_shifts = list(rule["shifts"][day_type].values())
                elif rule["rule_type"] == "fixed_time":
                    available_shifts = [rule["default_shift"]]
                    if "special_duty" in rule:
                        special_duty = rule["special_duty"]
                        if special_duty["frequency"] == "monthly" and date.day == special_duty["day_of_month"]:
                            available_shifts.append(special_duty["shift"])
            
            # Define role-based colors (you can adjust these as needed)
            role_colors = {
                "Freelancer": (75, 150, 225),      # Light blue
                "SeniorEditor": (225, 75, 75),     # Red
                "economics": (75, 225, 75),        # Green
                "Entertainment": (225, 225, 75),   # Yellow
                "KoreanEntertainment": (225, 75, 225)  # Purple
            }
            
            # Get color for the current role or use default light blue
            color_rgb = role_colors.get(role, (75, 150, 225))  # Default to light blue if role not found
            color = QColor(*color_rgb)
            
            for shift in available_shifts:
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
        filtered_employees = [emp.name for emp in self.employees if role == "All" or emp.employee_type == role]

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
        old_employee = next((emp for emp in self.employees if emp.name == name), None)
        
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
        self.availability = load_data()  # Reload availability data
        self.update_employee_list(self.role_combo.currentText())
        self.update_calendar() # Refresh the calendar
        dialog.accept()


    def save_data(self):
        sync_availability()  # Force synchronization
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
            # Generate schedule without exporting to Excel
            warnings = generate_schedule(self.availability, self.start_date, export_to_excel=False)
            
            # Get the schedule data from the function
            # You need to modify your scheduling_logic.py to return the schedule data
            from scheduling_logic import get_last_generated_schedule
            schedule_data = get_last_generated_schedule()
            
            if warnings:
                warning_message = "\n".join(warnings)
                QMessageBox.warning(self, "Scheduling Warnings", warning_message)
            
            # Show preview dialog
            preview_dialog = SchedulePreviewDialog(schedule_data, self)
            preview_dialog.exec_()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate schedule: {str(e)}")


    def import_from_google_form(self):
        """Open file dialog to select and import Google Form response Excel file."""
        from PySide6.QtWidgets import QFileDialog
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Google Form Response Excel File",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            try:
                from scheduling_logic import import_from_google_form
                result = import_from_google_form(file_path)
                
                # Reload data and update UI
                self.availability = load_data()
                self.employees = load_employees()
                self.employee_names = [emp.name for emp in self.employees if isinstance(emp, Freelancer)]
                
                # Update UI components
                self.update_employee_list(self.role_combo.currentText())
                self.update_calendar()
                
                QMessageBox.information(self, "Success", result)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to import Google Form data: {str(e)}")

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
            
            # Refresh data and UI components
            self.employees = load_employees()
            self.availability = load_data()
            
            self.update_employee_list(self.role_combo.currentText())  # Refresh employee list in UI
            self.update_calendar()  # Refresh calendar view
            
            QMessageBox.information(self, "Success", f"Employee {name} deleted successfully.")


class SchedulePreviewDialog(QDialog):
    def __init__(self, schedule_data, parent=None):
        super().__init__(parent)
        self.schedule_data = schedule_data
        self.setWindowTitle("Schedule Preview")
        self.setGeometry(200, 200, 800, 500)
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        # Create table widget
        self.table = QTableWidget()
        if self.schedule_data:
            # Set row and column count
            self.table.setRowCount(len(self.schedule_data))
            self.table.setColumnCount(len(self.schedule_data[0]))
            
            # Set headers
            headers = list(self.schedule_data[0].keys())
            self.table.setHorizontalHeaderLabels(headers)
            
            # Populate table
            for row, entry in enumerate(self.schedule_data):
                for col, header in enumerate(headers):
                    item = QTableWidgetItem(str(entry.get(header, "")))
                    self.table.setItem(row, col, item)
            
            # Resize columns to content
            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        
        layout.addWidget(self.table)
        
        # Add buttons
        button_box = QDialogButtonBox()
        export_btn = QPushButton("導出")
        export_btn.clicked.connect(self.export_to_excel)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        
        button_box.addButton(export_btn, QDialogButtonBox.AcceptRole)
        button_box.addButton(cancel_btn, QDialogButtonBox.RejectRole)
        
        layout.addWidget(button_box)
    
    def export_to_excel(self):
        try:
            df = pd.DataFrame(self.schedule_data)
            df.to_excel("schedule_with_senior_editors.xlsx", index=False)
            QMessageBox.information(self, "Success", "Excel schedule has been successfully generated!")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export schedule: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AvailabilityEditor()
    window.show()
    sys.exit(app.exec())
