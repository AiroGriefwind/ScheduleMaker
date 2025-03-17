#Core Components and Libraries
import sys
import json
import pandas as pd
from collections import deque  
from datetime import datetime, timedelta
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                              QHBoxLayout, QPushButton, QComboBox, QLabel,
                              QMessageBox, QGridLayout, QScrollArea)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor, QPalette

# Create a base Employee class with common functionality
class Employee:
    def __init__(self, name, employee_type):
        self.name = name
        self.employee_type = employee_type
        
    def get_available_shifts(self):
        # Base implementation for available shifts
        pass
    
    def get_display_name(self):
        return f"{self.name}({self.employee_type[0]})"
        
# Define specialized employee classes
class Freelancer(Employee):
    def __init__(self, name):
        super().__init__(name, "Freelancer")
        
    def get_available_shifts(self):
        # Return freelancer-specific shifts
        return ["7-16", "10-19", "15-24"]
    
# Initialize employees
def init_employees():
    freelancers = [
        Freelancer("Helen"),
        Freelancer("Lili"),
        Freelancer("Matthew"),
        Freelancer("Ka"),
        Freelancer("Kit"),
        Freelancer("Paul")
    ]
    return freelancers
    
# Initialize availability data structure
def init_availability(start_date, employees):
    return {
        (start_date + timedelta(days=i)).strftime("%Y-%m-%d"): {
            employee.get_display_name(): [] for employee in employees
        } for i in range(7)
    }

# Load/save data
def save_data(data):
    with open('availability.json', 'w') as f:
        json.dump(data, f)

def load_data():
    try:
        with open('availability.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return None

# ---------- Constants ----------
# Initialize employees
EMPLOYEES = init_employees()
# For backward compatibility
FREELANCERS = [employee.get_display_name() for employee in EMPLOYEES if isinstance(employee, Freelancer)]
# Shift colors for UI
SHIFT_COLORS = {
    "7-16": QColor(144, 238, 144),   # Light green
    "10-19": QColor(255, 228, 181),  # Light orange
    "15-24": QColor(176, 224, 230)   # Light blue
}
RULES = {
    "weekday": {"early": "7-16", "day": "10-19", "night": "15-24"},
    "weekend": {"early": "7-16", "day": "10-19", "night": "15-24"}
}

class AvailabilityEditor(QMainWindow):
    def __init__(self, start_date=datetime(2025, 3, 17)):
        super().__init__()
        self.start_date = start_date
        self.employees = EMPLOYEES
        self.employee_names = [employee.get_display_name() for employee in self.employees]
        self.current_employee_name = self.employee_names[0] if self.employee_names else ""
        
        # Check if we have existing data or initialize new
        loaded_data = load_data()
        if loaded_data:
            self.availability = loaded_data
        else:
            self.availability = init_availability(start_date, self.employees)
        
        self.init_ui()
        self.update_calendar()

    def init_ui(self):
        self.setWindowTitle("Employee Availability Editor")
        self.setGeometry(100, 100, 1000, 600)
        
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # Top controls
        control_layout = QHBoxLayout()
        self.employee_combo = QComboBox()
        self.employee_combo.addItems(self.employee_names)
        self.employee_combo.currentTextChanged.connect(self.employee_changed)
        control_layout.addWidget(QLabel("Select Employee:"))
        control_layout.addWidget(self.employee_combo)
        
        # Calendar scroll area
        scroll = QScrollArea()
        self.calendar_widget = QWidget()
        self.calendar_layout = QGridLayout(self.calendar_widget)
        scroll.setWidget(self.calendar_widget)
        scroll.setWidgetResizable(True)
        
        # Bottom buttons
        button_layout = QHBoxLayout()
        save_btn = QPushButton("保存時間")
        save_btn.clicked.connect(self.save_data)
        generate_btn = QPushButton("導出至Excel")
        generate_btn.clicked.connect(self.generate_schedule)
        button_layout.addWidget(save_btn)
        button_layout.addWidget(generate_btn)
        
        layout.addLayout(control_layout)
        layout.addWidget(scroll)
        layout.addLayout(button_layout)

        import_btn = QPushButton("從Excel導入")
        import_btn.clicked.connect(lambda: self.import_from_excel("availability_export.xlsx"))  # Replace with file dialog if needed
        button_layout.addWidget(import_btn)

        export_btn = QPushButton("導出時間表至Excel")
        export_btn.clicked.connect(self.export_availability_to_excel)
        button_layout.addWidget(export_btn)

        clear_btn = QPushButton("Clear Availability")
        clear_btn.clicked.connect(self.clear_availability)
        button_layout.addWidget(clear_btn)


    def clear_availability(self):
        # Confirm with the user before clearing
        reply = QMessageBox.question(
            self,
            "Confirm Clear",
            "Are you sure you want to clear all stored availability?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            # Reset availability
            self.availability = init_availability(self.start_date, self.employees)
            save_data(self.availability)
            
            # Update calendar UI
            self.update_calendar()
            
            QMessageBox.information(self, "Cleared", "All availability data has been cleared!")


    def import_from_excel(self, file_path):
        try:
            # Read Excel file into a DataFrame
            df = pd.read_excel(file_path)
            
            # Validate required columns
            required_columns = {'Date', 'Employee', 'Shift'}
            if not required_columns.issubset(df.columns):
                raise ValueError(f"Excel file must contain columns: {required_columns}")
            
            # Reset availability based on imported data
            self.availability = {}
            for _, row in df.iterrows():
                date_str = row['Date']
                employee_name = row['Employee']
                shift = row['Shift']
                
                if date_str not in self.availability:
                    self.availability[date_str] = {name: [] for name in self.employee_names}
                
                if employee_name not in self.availability[date_str]:
                    self.availability[date_str][employee_name] = []
                
                if shift not in self.availability[date_str][employee_name]:
                    self.availability[date_str][employee_name].append(shift)
            
            # Save updated availability and refresh UI
            save_data(self.availability)
            self.update_calendar()
            
            QMessageBox.information(self, "Success", "Data imported successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to import data: {str(e)}")

    
    def export_availability_to_excel(self):
        try:
            # Prepare data for export
            data = []
            for date, employees in self.availability.items():
                for employee_name, shifts in employees.items():
                    for shift in shifts:
                        data.append({"Date": date, "Employee": employee_name, "Shift": shift})
            
            # Convert to DataFrame
            df = pd.DataFrame(data)
            
            # Save to Excel
            df.to_excel("availability_export.xlsx", index=False)
            
            QMessageBox.information(self, "Success", "Availability exported successfully to Excel!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export availability: {str(e)}")


    def create_day_widget(self, date_str):
        day_widget = QWidget()
        day_layout = QVBoxLayout(day_widget)
        
        date = datetime.strptime(date_str, "%Y-%m-%d")
        day_label = QLabel(date.strftime("%a\n%d %b"))
        day_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        day_layout.addWidget(day_label)
        
         # Get the employee object for the current employee name
        current_employee = next((e for e in self.employees if e.get_display_name() == self.current_employee_name), None)
        
        if current_employee:
            # Get shifts based on employee type
            available_shifts = current_employee.get_available_shifts()
            
            for shift in available_shifts:
                if shift in SHIFT_COLORS:
                    color = SHIFT_COLORS[shift]
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
        # Clear existing calendar
        for i in reversed(range(self.calendar_layout.count())): 
            self.calendar_layout.itemAt(i).widget().setParent(None)
        
        # Create new calendar days
        sorted_dates = sorted(self.availability.keys())
        for col, date_str in enumerate(sorted_dates):
            day_widget = self.create_day_widget(date_str)
            self.calendar_layout.addWidget(day_widget, 0, col)
            
            # Update button states for each shift
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
            
            # Save changes and refresh UI
            save_data(self.availability)
            self.update_calendar()


    def employee_changed(self, name):
        self.current_employee_name = name
        self.update_calendar()

    def save_data(self):
        save_data(self.availability)
        QMessageBox.information(self, "Saved", "時間已成功保存!")

    # Generates work schedule based on availability and shift requirements
    # EXTENSION POINT: Should use different scheduling rules per employee type
    def generate_schedule(self):
        try:
            date_strings = sorted(self.availability.keys())
            dates = [datetime.strptime(d, "%Y-%m-%d") for d in date_strings]
            schedule = []

            # Get only freelancers for now
            freelancer_names = [e.get_display_name() for e in self.employees if isinstance(e, Freelancer)]
            
            # Shift staffing requirements - should be configurable per employee type
            SHIFT_REQUIREMENTS = {
                "weekday": {"early": 1, "day": 1, "night": 2},
                "weekend": {"early": 1, "day": 1, "night": 1}
            }
            
            # Mapping between shift names and time ranges
            SHIFT_MAPPING = {
                "early": "7-16",
                "day": "10-19",
                "night": "15-24"
            }

            # Process each date in the schedule period
            for date in dates:
                # Determine if date is weekday or weekend
                day_type = 'weekend' if date.weekday() >= 5 else 'weekday'
                # Initialize all employees as off-duty
                assigned_shifts = {name: 'off' for name in FREELANCERS}
                # Track available employees for each shift
                available_freelancers = {shift: [] for shift in SHIFT_MAPPING.values()}
                
                # Populate available employees list based on their availability
                for freelancer_name in freelancer_names:
                    for shift, shift_time in SHIFT_MAPPING.items():
                        if shift_time in self.availability[date.strftime("%Y-%m-%d")][freelancer_name]:
                            available_freelancers[shift_time].append(freelancer_name)
                
                # Assign shifts based on staffing requirements
                for shift, count in SHIFT_REQUIREMENTS[day_type].items():
                    shift_time = SHIFT_MAPPING[shift]
                    assigned = 0
                    while assigned < count and available_freelancers[shift_time]:
                        for freelancer_name in available_freelancers[shift_time]:
                            if assigned_shifts[freelancer_name] == 'off':
                                assigned_shifts[freelancer_name] = shift_time
                                assigned += 1
                                # Remove assigned employee from available lists
                                for s in SHIFT_MAPPING.values():
                                    if freelancer_name in available_freelancers[s]:
                                        available_freelancers[s].remove(freelancer_name)
                                break
                        if assigned == count:
                            break
                
                # Add completed day schedule to result
                schedule.append({"Date": date.strftime("%d/%m/%Y"), **assigned_shifts})

            # Save schedule to Excel
            df = pd.DataFrame(schedule)
            df.to_excel("freelancer_schedule.xlsx", index=False)
            QMessageBox.information(self, "Success", "Excel已成功生成!")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AvailabilityEditor()
    window.show()
    sys.exit(app.exec())
