import sys
import json
import pandas as pd
from collections import deque  # <-- ADD THIS IMPORT
from datetime import datetime, timedelta
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                              QHBoxLayout, QPushButton, QComboBox, QLabel,
                              QMessageBox, QGridLayout, QScrollArea)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor, QPalette

# Initialize availability data structure
def init_availability(start_date):
    return {
        (start_date + timedelta(days=i)).strftime("%Y-%m-%d"): {
            freelancer: [] for freelancer in FREELANCERS
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
FREELANCERS = ["Helen(F)", "Lili(F)", "Matthew(F)", "Ka(F)", "Kit(F)", "Paul(F)"]
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
    def __init__(self, start_date=datetime(2025, 2, 1)):
        super().__init__()
        self.start_date = start_date
        self.current_freelancer = FREELANCERS[0]
        self.availability = load_data() or init_availability(start_date)
        
        self.init_ui()
        self.update_calendar()

    def init_ui(self):
        self.setWindowTitle("Freelancer Availability Editor")
        self.setGeometry(100, 100, 1000, 600)
        
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # Top controls
        control_layout = QHBoxLayout()
        self.freelancer_combo = QComboBox()
        self.freelancer_combo.addItems(FREELANCERS)
        self.freelancer_combo.currentTextChanged.connect(self.freelancer_changed)
        control_layout.addWidget(QLabel("Select Freelancer:"))
        control_layout.addWidget(self.freelancer_combo)
        
        # Calendar scroll area
        scroll = QScrollArea()
        self.calendar_widget = QWidget()
        self.calendar_layout = QGridLayout(self.calendar_widget)
        scroll.setWidget(self.calendar_widget)
        scroll.setWidgetResizable(True)
        
        # Bottom buttons
        button_layout = QHBoxLayout()
        save_btn = QPushButton("Save Availability")
        save_btn.clicked.connect(self.save_data)
        generate_btn = QPushButton("Generate Schedule")
        generate_btn.clicked.connect(self.generate_schedule)
        button_layout.addWidget(save_btn)
        button_layout.addWidget(generate_btn)
        
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
        
        for shift, color in SHIFT_COLORS.items():
            btn = QPushButton(shift)
            btn.setCheckable(True)
            btn.setStyleSheet(f"""
                QPushButton {{ 
                    background-color: {color.name()}; 
                    border: 1px solid gray;
                }}
                QPushButton:checked {{
                    border: 2px solid black;
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
            
            # Update button states
            shifts = self.availability[date_str][self.current_freelancer]
            for i in range(day_widget.layout().count()):
                widget = day_widget.layout().itemAt(i).widget()
                if isinstance(widget, QPushButton):
                    widget.setChecked(widget.text() in shifts)

    def toggle_shift(self, date_str, shift):
        current_shifts = self.availability[date_str][self.current_freelancer]
        if shift in current_shifts:
            current_shifts.remove(shift)
        else:
            current_shifts.append(shift)
        self.update_calendar()

    def freelancer_changed(self, name):
        self.current_freelancer = name
        self.update_calendar()

    def save_data(self):
        save_data(self.availability)
        QMessageBox.information(self, "Saved", "Availability data saved successfully!")

    def generate_schedule(self):
        try:
            date_strings = sorted(self.availability.keys())
            dates = [datetime.strptime(d, "%Y-%m-%d") for d in date_strings]
            schedule = []
            
            # Updated scheduling rules with shift priorities
            SHIFT_PRIORITY = {
                "weekday": [
                    ("night", "15-24", 2),  # Require 2 night shifts first
                    ("early", "7-16", 1),
                    ("day", "10-19", 1)
                ],
                "weekend": [
                    ("early", "7-16", 1),
                    ("day", "10-19", 1),
                    ("night", "15-24", 1)
                ]
            }

            for date in dates:
                day_type = 'weekend' if date.weekday() >= 5 else 'weekday'
                assigned_shifts = {name: 'off' for name in FREELANCERS}
                freelancer_queue = deque(FREELANCERS)  # Reset queue daily

                # Process shifts in priority order
                for shift_name, shift_value, required in SHIFT_PRIORITY[day_type]:
                    assignments_made = 0
                    rotation_count = 0

                    while assignments_made < required and rotation_count < len(FREELANCERS):
                        current_worker = freelancer_queue[0]
                        
                        if (assigned_shifts[current_worker] == 'off' and 
                            shift_value in self.availability[date.strftime("%Y-%m-%d")][current_worker]):
                            assigned_shifts[current_worker] = shift_value
                            freelancer_queue.rotate(-1)  # Move to next worker
                            assignments_made += 1
                            rotation_count = 0  # Reset rotation counter
                        else:
                            freelancer_queue.rotate(-1)
                            rotation_count += 1

                schedule.append({"Date": date.strftime("%d/%m/%Y"), **assigned_shifts})

            # Save to Excel
            df = pd.DataFrame(schedule)
            df.to_excel("freelancer_schedule.xlsx", index=False)
            QMessageBox.information(self, "Success", "Schedule generated successfully!")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AvailabilityEditor()
    window.show()
    sys.exit(app.exec())
