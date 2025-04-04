import sys
import pandas as pd # Import for SchedulePreviewDialog
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                              QHBoxLayout, QPushButton, QComboBox, QLabel,
                              QMessageBox, QGridLayout, QScrollArea, QDialog, QLineEdit, QMenu,
                                )
from PySide6.QtWidgets import QTableWidget, QTableWidgetItem, QHeaderView, QDialogButtonBox # Import for SchedulePreviewDialog
from PySide6.QtCore import Qt, QDate, QPoint
from PySide6.QtGui import QColor, QPalette
from PySide6.QtGui import QCursor # Import for correctly positioning leave menu
from scheduling_logic import (EMPLOYEES, Freelancer,  
                              load_data, save_data, init_availability, 
                               generate_schedule, import_from_excel, 
                               edit_employee, load_employees, ROLE_RULES, add_employee, delete_employee,sync_availability,
                              export_availability_to_excel, clear_availability)
# Import for Calendar UI
from datetime import datetime, timedelta 

# Import for logging
from logger_utils import setup_logging, log_error, log_info, create_data_package  

# Import for localization
# after updating the .ts file, run this in console to generate the .qm file
#& "C:/Users/wangz/AppData/Local/Programs/Python/Python313/Lib/site-packages/PySide6/lrelease.exe" zh_TW.ts
from PySide6.QtCore import QTranslator, QFile, QIODevice
from PySide6.QtXml import QDomDocument

def compile_ts_to_qm(ts_file, qm_file):
    """Simple function to convert .ts file content to .qm format"""
    try:
        # This is a simplified version - in practice, use proper tools
        translator = QTranslator()
        if translator.load(ts_file):
            return translator.save(qm_file)
        return False
    except Exception as e:
        print(f"Error compiling translation: {e}")
        return False

# Try to compile the .ts file
compile_ts_to_qm("zh_TW.ts", "zh_TW.qm")


class AvailabilityEditor(QMainWindow):
    def __init__(self, start_date=datetime(2025, 3, 17)):
        super().__init__()

        # Initialize logging first thing
        self.log_file = setup_logging()
        log_info("Application started")

        self.start_date = start_date
        self.employees = load_employees()  # Load employees from JSON
        self.employee_names = [emp.name for emp in self.employees if isinstance(emp, Freelancer)]
        self.current_employee_name = self.employee_names[0] if self.employee_names else ""
        
        loaded_data = load_data()
        self.availability = loaded_data if loaded_data else init_availability(start_date, self.employees)
        
        self.init_ui()
        self.update_calendar()

    def init_ui(self):
        self.setWindowTitle(self.tr("Employee Availability Editor"))
        self.setGeometry(100, 100, 1000, 600)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        # Create menu bar at the top instead of bottom
        menu_bar_layout = QHBoxLayout()

        # FILE MENU
        file_btn = QPushButton(self.tr("File"))
        file_menu = QMenu(self)
        save_action = file_menu.addAction(self.tr("Save Availability"))
        file_menu.addSeparator()
        import_excel_action = file_menu.addAction(self.tr("Import from Excel"))
        import_form_action = file_menu.addAction(self.tr("Import from Google Form"))
        export_avail_action = file_menu.addAction(self.tr("Export Availability to Excel"))
        file_menu.addSeparator()
        clear_action = file_menu.addAction(self.tr("Clear Availability"))
        file_btn.setMenu(file_menu)

        # EDIT MENU
        edit_btn = QPushButton(self.tr("Edit"))
        edit_menu = QMenu(self)
        add_employee_action = edit_menu.addAction(self.tr("Add Employee"))
        add_role_action = edit_menu.addAction(self.tr("Add Role"))
        edit_btn.setMenu(edit_menu)

        # TOOLS MENU
        tools_btn = QPushButton(self.tr("Tools"))
        tools_menu = QMenu(self)
        generate_action = tools_menu.addAction(self.tr("Generate Schedule"))
        validate_action = tools_menu.addAction(self.tr("Validate Schedule"))
        tools_btn.setMenu(tools_menu)

        # HELP MENU
        help_btn = QPushButton(self.tr("Help"))
        help_menu = QMenu(self)
        data_package_action = help_menu.addAction(self.tr("Create Data Package"))
        help_btn.setMenu(help_menu)

        # Add all menu buttons to the layout
        menu_bar_layout.addWidget(file_btn)
        menu_bar_layout.addWidget(edit_btn)
        menu_bar_layout.addWidget(tools_btn)
        menu_bar_layout.addWidget(help_btn)
        menu_bar_layout.addStretch()  # Push buttons to the left

        # Connect all actions to their respective functions
        save_action.triggered.connect(self.save_data)
        import_excel_action.triggered.connect(self.import_from_excel)
        import_form_action.triggered.connect(self.import_from_google_form)
        export_avail_action.triggered.connect(self.export_availability_to_excel)
        clear_action.triggered.connect(self.clear_availability)
        add_employee_action.triggered.connect(self.add_new_employee)
        add_role_action.triggered.connect(self.add_new_role)
        generate_action.triggered.connect(self.generate_schedule)
        validate_action.triggered.connect(self.validate_schedule)
        data_package_action.triggered.connect(self.create_debug_package)

        # Style the menu buttons to look more like Google Docs
        menu_button_style = """
        QPushButton {
            background-color: transparent;
            border: none;
            padding: 8px 16px;
            text-align: center;
            font-size: 14px;
        }
        QPushButton:hover {
            background-color: #e8e8e8;
        }
        QPushButton:pressed {
            background-color: #d0d0d0;
        }
        QPushButton::menu-indicator {
            width: 0px;
        }
        """

        file_btn.setStyleSheet(menu_button_style)
        edit_btn.setStyleSheet(menu_button_style)
        tools_btn.setStyleSheet(menu_button_style)
        help_btn.setStyleSheet(menu_button_style)

        # Add the menu bar to the main layout
        layout.addLayout(menu_bar_layout)

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
        control_layout.addWidget(QLabel(self.tr("Select Role:")))
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
        control_layout.addWidget(QLabel(self.tr("Select Employee:")))
        control_layout.addWidget(employee_scroll)

        layout.addLayout(control_layout)

        # Calendar container modifications
        calendar_container = QVBoxLayout()
        self.selected_employee_label = QLabel(self.tr("Currently Selected: ") + self.current_employee_name)
        self.selected_employee_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        calendar_container.addWidget(self.selected_employee_label)
        
        # Scroll area for the calendar - adjust to better handle grid layout
        scroll = QScrollArea()
        self.calendar_widget = QWidget()
        self.calendar_layout = QGridLayout(self.calendar_widget)
        self.calendar_layout.setSpacing(10)  # Add spacing between cells
        self.calendar_layout.setContentsMargins(10, 10, 10, 10)  # Add margins
        
        scroll.setWidget(self.calendar_widget)
        scroll.setWidgetResizable(True)
        
        # Add scroll area to calendar container
        calendar_container.addWidget(scroll)
        
        # Add calendar container to main layout
        layout.addLayout(calendar_container)

        def setup_menu_styles(self):
            # Apply styles to all dropdown menus
            menu_style = """
            QMenu {
                background-color: white;
                border: 1px solid #c0c0c0;
                padding: 5px 0px;
            }
            QMenu::item {
                padding: 6px 25px 6px 20px;
                border: 1px solid transparent;
            }
            QMenu::item:selected {
                background-color: #e8e8e8;
            }
            """
            
            # Apply the style to all menus
            for menu in [file_menu, edit_menu, tools_menu, help_menu]:
                menu.setStyleSheet(menu_style)


    def add_new_role(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add New Role")
        
        layout = QVBoxLayout(dialog)
        
        # Role name input
        layout.addWidget(QLabel("Role Name:"))
        role_name_input = QLineEdit()
        layout.addWidget(role_name_input)
        
        # Rule type selection
        layout.addWidget(QLabel("Rule Type:"))
        rule_type_combo = QComboBox()
        rule_type_combo.addItems(["fixed_time", "shift_based"])
        layout.addWidget(rule_type_combo)
        
        # Default shift for fixed_time roles
        layout.addWidget(QLabel("Default Shift (e.g., 10-19):"))
        default_shift_input = QLineEdit()
        layout.addWidget(default_shift_input)
        
        # Container for shift-based settings (initially hidden)
        shift_based_container = QWidget()
        shift_layout = QVBoxLayout(shift_based_container)
        
        # Weekday shifts
        shift_layout.addWidget(QLabel("Weekday Shifts (comma-separated, e.g., 7-16,10-19,15-24):"))
        weekday_shifts_input = QLineEdit()
        shift_layout.addWidget(weekday_shifts_input)
        
        # Weekend shifts
        shift_layout.addWidget(QLabel("Weekend Shifts (comma-separated, e.g., 7-16,10-19,15-24):"))
        weekend_shifts_input = QLineEdit()
        shift_layout.addWidget(weekend_shifts_input)
        
        layout.addWidget(shift_based_container)
        shift_based_container.setVisible(False)
        
        # Show/hide shift-based settings based on rule type
        def on_rule_type_changed(rule_type):
            is_shift_based = (rule_type == "shift_based")
            shift_based_container.setVisible(is_shift_based)
            default_shift_input.setEnabled(not is_shift_based)
        
        rule_type_combo.currentTextChanged.connect(on_rule_type_changed)
        
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self.save_new_role(
            dialog, 
            role_name_input.text(),
            rule_type_combo.currentText(),
            default_shift_input.text(),
            weekday_shifts_input.text(),
            weekend_shifts_input.text()
        ))
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        dialog.exec_()

    def save_new_role(self, dialog, role_name, rule_type, default_shift, weekday_shifts, weekend_shifts):
        if not role_name:
            log_error("Role name cannot be empty")
            QMessageBox.warning(self, "Error", "Role name cannot be empty")
            return
        
        if role_name in ROLE_RULES:
            QMessageBox.warning(self, "Error", f"Role '{role_name}' already exists")
            return
        
        if rule_type == "fixed_time" and not default_shift:
            QMessageBox.warning(self, "Error", "Default shift is required for fixed_time roles")
            return
        
        if rule_type == "shift_based" and (not weekday_shifts or not weekend_shifts):
            QMessageBox.warning(self, "Error", "Weekday and weekend shifts are required for shift_based roles")
            return
        
        # Create new role configuration
        new_role = {"rule_type": rule_type}
        
        if rule_type == "fixed_time":
            new_role["default_shift"] = default_shift
        else:  # shift_based
            weekday_shift_list = [s.strip() for s in weekday_shifts.split(",")]
            weekend_shift_list = [s.strip() for s in weekend_shifts.split(",")]
            
            weekday_shift_dict = {f"shift{i+1}": shift for i, shift in enumerate(weekday_shift_list)}
            weekend_shift_dict = {f"shift{i+1}": shift for i, shift in enumerate(weekend_shift_list)}
            
            new_role["shifts"] = {
                "weekday": weekday_shift_dict,
                "weekend": weekend_shift_dict
            }
            
            # Default requirements (can be enhanced in future versions)
            new_role["requirements"] = {
                "weekday": {f"shift{i+1}": 1 for i in range(len(weekday_shift_list))},
                "weekend": {f"shift{i+1}": 1 for i in range(len(weekend_shift_list))}
            }
        
        # Add the new role to ROLE_RULES
        from scheduling_logic import add_role
        add_role(role_name, new_role)
        
        # Update UI components
        self.role_combo.clear()
        self.role_combo.addItems(["All"] + list(ROLE_RULES.keys()))
        
        QMessageBox.information(self, "Success", f"Role '{role_name}' added successfully")
        dialog.accept()


    def add_new_employee(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add New Employee")
        
        layout = QVBoxLayout(dialog)
        
        name_input = QLineEdit()
        
        # Role section with add/remove functionality
        role_section = QWidget()
        role_layout = QVBoxLayout(role_section)
        role_layout.setContentsMargins(0, 0, 0, 0)
        
        # Primary role (Role1)
        primary_role_layout = QHBoxLayout()
        primary_role_label = QLabel("Role1 (Primary):")
        primary_role_input = QComboBox()
        primary_role_input.addItems(list(ROLE_RULES.keys()))
        primary_role_layout.addWidget(primary_role_label)
        primary_role_layout.addWidget(primary_role_input)
        
        # Add button for additional roles
        add_role_btn = QPushButton("+")
        add_role_btn.setFixedWidth(30)
        primary_role_layout.addWidget(add_role_btn)
        
        role_layout.addLayout(primary_role_layout)
        
        # Container for additional roles
        additional_roles_container = QWidget()
        additional_roles_layout = QVBoxLayout(additional_roles_container)
        additional_roles_layout.setContentsMargins(0, 0, 0, 0)
        role_layout.addWidget(additional_roles_container)
        
        # List to track additional role inputs
        additional_role_inputs = []
        
        # Function to add a new role input
        def add_role_input():
            role_row = QHBoxLayout()
            role_label = QLabel(f"Role{len(additional_role_inputs) + 2}:")
            role_combo = QComboBox()
            role_combo.addItems(list(ROLE_RULES.keys()))
            
            remove_btn = QPushButton("-")
            remove_btn.setFixedWidth(30)
            
            role_row.addWidget(role_label)
            role_row.addWidget(role_combo)
            role_row.addWidget(remove_btn)
            
            # Create a container for this row
            row_container = QWidget()
            row_container.setLayout(role_row)
            
            additional_roles_layout.addWidget(row_container)
            additional_role_inputs.append((row_container, role_combo))
            
            # Connect remove button
            remove_btn.clicked.connect(lambda: remove_role_input(row_container))
        
        # Function to remove a role input
        def remove_role_input(container):
            # Find and remove from tracking list
            for i, (cont, combo) in enumerate(additional_role_inputs):
                if cont == container:
                    additional_role_inputs.pop(i)
                    break
            
            # Remove from UI
            container.setParent(None)
            
            # Update labels for remaining roles
            for i, (cont, combo) in enumerate(additional_role_inputs):
                cont.layout().itemAt(0).widget().setText(f"Role{i + 2}:")
        
        # Connect add button
        add_role_btn.clicked.connect(add_role_input)
        
        # Time inputs
        start_time_input = QLineEdit()
        end_time_input = QLineEdit()
        
        # Build the form
        layout.addWidget(QLabel("Name:"))
        layout.addWidget(name_input)
        layout.addWidget(role_section)
        layout.addWidget(QLabel("Start Time (HH:MM):"))
        layout.addWidget(start_time_input)
        layout.addWidget(QLabel("End Time (HH:MM):"))
        layout.addWidget(end_time_input)
        
        # Add role change listener for primary role
        def primary_role_changed(role):
            is_freelancer = (role == "Freelancer")
            start_time_input.setEnabled(not is_freelancer)
            end_time_input.setEnabled(not is_freelancer)
        
        primary_role_input.currentTextChanged.connect(primary_role_changed)
        primary_role_changed(primary_role_input.currentText())  # Initial check
        
        # Save button
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(lambda: self.save_new_employee(
            dialog, 
            name_input.text(), 
            primary_role_input.currentText(),
            [combo.currentText() for _, combo in additional_role_inputs],
            start_time_input.text(), 
            end_time_input.text()
        ))
        
        layout.addWidget(save_btn)
        dialog.exec_()

    def save_new_employee(self, dialog, new_name, new_role, additional_roles=None, new_start_time=None, new_end_time=None):
        if new_role != 'Freelancer' and (not new_start_time or not new_end_time):
            QMessageBox.warning(self, "Error", "Fulltimers must provide start and end times")
            return
        
        # For freelancers, ignore time inputs and use default rules
        if new_role == 'Freelancer':
            add_employee(new_name, new_role, additional_roles)
        else:
            add_employee(new_name, new_role, additional_roles, new_start_time, new_end_time)
        
        # Refresh data and UI
        self.employees = load_employees()
        self.availability = load_data()
        self.update_employee_list(self.role_combo.currentText())
        self.update_calendar()
        dialog.accept()


    # Updated ui.py
    def save_new_employee(self, dialog, new_name, new_role, new_start_time=None, new_end_time=None):
        if new_role != 'Freelancer' and (not new_start_time or not new_end_time):
            QMessageBox.warning(self, "Error", "Fulltimers must provide start and end times")
            return
        
        # For freelancers, ignore time inputs and use default rules
        if new_role == 'Freelancer':
            add_employee(new_name, new_role)
        else:
            add_employee(new_name, new_role, new_start_time, new_end_time)
        
        # Refresh data and UI
        self.employees = load_employees()
        self.availability = load_data()
        self.update_employee_list(self.role_combo.currentText())
        self.update_calendar()
        dialog.accept()

    def show_shift_context_menu(self, pos, date_str, shift):
        
        
        context_menu = QMenu()
        leave_action = context_menu.addAction("Leave Application")

        # Execute the menu at the current cursor position
        action = context_menu.exec_(QCursor.pos())
        
        if action == leave_action:
            self.show_leave_dialog(date_str, shift)


    def show_leave_dialog(self, date_str, shift):
        dialog = LeaveDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            leave_type = dialog.get_leave_type()
            self.set_leave(date_str, shift, leave_type)

    def set_leave(self, date_str, shift, leave_type):
        if self.current_employee_name in self.availability[date_str]:
            current_shifts = self.availability[date_str][self.current_employee_name]
            if shift in current_shifts:
                current_shifts.remove(shift)
            current_shifts.append(leave_type)
            
            save_data(self.availability)
            self.update_calendar()



    def create_day_widget(self, date_str):
        day_widget = QWidget()
        day_layout = QVBoxLayout(day_widget)
        day_layout.setContentsMargins(5, 5, 5, 5)  # Add some padding
        
        # Date label setup with better formatting
        date = datetime.strptime(date_str, "%Y-%m-%d")
        day_label = QLabel(date.strftime("%a\n%d %b"))
        day_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Change the styling to improve contrast
        day_label.setStyleSheet("""
            font-weight: bold; 
            background-color: #404040; 
            color: #ffffff;
            padding: 5px;
            border-radius: 3px;
        """)

        day_layout.addWidget(day_label)

        current_employee = next((e for e in self.employees if e.name == self.current_employee_name), None)

        if current_employee:
            role = current_employee.employee_type
            rule = ROLE_RULES.get(role, {})
            available_shifts = []

            # Determine available shifts based on role rules
            if rule:
                if rule["rule_type"] == "shift_based":
                    day_type = "weekend" if date.weekday() >= 5 else "weekday"
                    available_shifts = list(rule["shifts"][day_type].values())
                elif rule["rule_type"] == "fixed_time":
                    # Use employee's custom time if available, otherwise use default
                    if current_employee.start_time and current_employee.end_time:
                        available_shifts = [f"{current_employee.start_time}-{current_employee.end_time}"]
                    else:
                        available_shifts = [rule["default_shift"]]

            # Get role color
            role_colors = {
                "Freelancer": (75, 150, 225),
                "SeniorEditor": (225, 75, 75),
                "economics": (75, 225, 75),
                "Entertainment": (225, 225, 75),
                "KoreanEntertainment": (225, 75, 225)
            }
            color = QColor(*role_colors.get(role, (75, 150, 225)))

            # Create shift buttons with leave management
            for shift in available_shifts:
                btn = QPushButton(shift)
                btn.setCheckable(True)
                btn.setProperty("original_shift", shift)  # Store original shift text

                # Style configuration
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background-color: lightgray;  /* Default deselected state */
                        border: 1px solid gray;
                        color: black;
                        padding: 4px;
                    }}
                    QPushButton:hover {{
                        background-color: {color.name()}; /* Hover changes to role-specific color */
                        color: white;
                    }}
                    QPushButton:checked {{
                        background-color: {color.name()}; /* Checked state matches hover color */
                        border: 2px solid white;
                        color: black;
                    }}
                """)

                # Enable interaction for all employee types
                btn.clicked.connect(lambda _, d=date_str, s=shift: self.toggle_shift(d, s))
                btn.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
                btn.customContextMenuRequested.connect(
                    lambda pos, d=date_str, s=shift: self.show_shift_context_menu(pos, d, s)
                )

                day_layout.addWidget(btn)

            # Update button states from availability data
            if self.current_employee_name in self.availability.get(date_str, {}):
                recorded_shifts = self.availability[date_str][self.current_employee_name]

                for i in range(day_layout.count()):
                    widget = day_layout.itemAt(i).widget()
                    if isinstance(widget, QPushButton) and widget != day_label:  # Skip the date label
                        current_text = widget.text()
                        original_shift = widget.property("original_shift")

                        # Check for leave status
                        if current_text in recorded_shifts:
                            widget.setChecked(True)
                        else:
                            # Handle leave types
                            leaves = [lt for lt in ["AL", "CL", "PH", "ON", "自由調配", "half off"] if lt in recorded_shifts]
                            if leaves:
                                widget.setText(leaves[0])
                                widget.setChecked(True)
                            else:
                                widget.setChecked(False)
                                # Restore original shift text if needed
                                if original_shift and widget.text() != original_shift:
                                    widget.setText(original_shift)

        return day_widget




    def update_calendar(self):
        # Clear existing widgets
        while self.calendar_layout.count():
            item = self.calendar_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        sorted_dates = sorted(self.availability.keys())
        
        # Group dates by week (Sunday to Saturday)
        weeks = []
        current_week = []
        
        # Find the first date
        if sorted_dates:
            first_date = datetime.strptime(sorted_dates[0], "%Y-%m-%d")
            
            # If the first date is not a Sunday, add placeholder dates
            if first_date.weekday() != 6:  # 6 is Sunday in Python's weekday() (0 is Monday)
                # Calculate days to go back to reach the previous Sunday
                days_to_previous_sunday = (first_date.weekday() + 1) % 7
                
                # Create placeholder dates for the beginning of the first week
                for i in range(days_to_previous_sunday, 0, -1):
                    placeholder_date = (first_date - timedelta(days=i)).strftime("%Y-%m-%d")
                    # Add placeholder to availability if it doesn't exist
                    if placeholder_date not in self.availability:
                        self.availability[placeholder_date] = {emp.name: [] for emp in self.employees}
                    
                    # Add to sorted_dates
                    sorted_dates.insert(0, placeholder_date)
        
        # Now group by weeks with Sunday as first day
        for date_str in sorted_dates:
            date = datetime.strptime(date_str, "%Y-%m-%d")
            
            # If this is the first date or it's a Sunday, start a new week
            if not current_week or date.weekday() == 6:  # 6 is Sunday in Python
                if current_week:  # If we have a non-empty current week, add it to weeks
                    weeks.append(current_week)
                current_week = [date_str]
            else:
                current_week.append(date_str)
        
        # Add the last week if it's not empty
        if current_week:
            weeks.append(current_week)
        
        # Display each week as a row in the grid
        for row, week in enumerate(weeks):
            for col, date_str in enumerate(week):
                day_widget = self.create_day_widget(date_str)
                self.calendar_layout.addWidget(day_widget, row, col)
                
                # Update button states for the current employee
                if self.current_employee_name in self.availability[date_str]:
                    current_shifts = self.availability[date_str][self.current_employee_name]
                    leaves = [s for s in current_shifts if s in {"AL", "CL", "PH", "ON", "自由調配", "half off"}]
                    
                    # Get employee and determine role type
                    current_employee = next((e for e in self.employees if e.name == self.current_employee_name), None)
                    is_freelancer = current_employee and current_employee.employee_type == "Freelancer"
                    
                    for i in range(day_widget.layout().count()):
                        widget = day_widget.layout().itemAt(i).widget()
                        if isinstance(widget, QPushButton):
                            original_shift = widget.property("original_shift")
                            
                            if leaves:
                                # Handle leave types
                                widget.setText(leaves[0])
                                widget.setChecked(True)
                            elif is_freelancer:
                                # For freelancers: Only check buttons for shifts that are selected
                                if original_shift in current_shifts:
                                    widget.setChecked(True)
                                else:
                                    widget.setChecked(False)
                                # Always preserve original shift text
                                widget.setText(original_shift)
                            else:
                                # For fixed shift employees: Handle as before
                                actual_shift = current_shifts[0] if current_shifts else ""
                                if actual_shift and actual_shift not in ["AL", "CL", "PH", "ON", "自由調配", "half off"]:
                                    widget.setText(actual_shift)
                                    widget.setChecked(True)

        
        

    def toggle_shift(self, date_str, shift):
        if self.current_employee_name in self.availability[date_str]:
            current_shifts = self.availability[date_str][self.current_employee_name]
            
            # Check if this is a leave type being toggled
            leave_types = ["AL", "CL", "PH", "ON", "自由調配", "half off"]
            
            if shift in leave_types:
                # Remove any existing leave types
                current_shifts = [s for s in current_shifts if s not in leave_types]
                current_shifts.append(shift)
            else:
                # If the shift is already in current_shifts, remove it
                if shift in current_shifts:
                    current_shifts.remove(shift)
                else:
                    # Remove any existing leave types before adding the new shift
                    current_shifts = [s for s in current_shifts if s not in leave_types]
                    current_shifts.append(shift)
            
            # Update the availability for this date and employee
            self.availability[date_str][self.current_employee_name] = current_shifts
            
            save_data(self.availability)
            self.update_calendar()





    def role_changed(self, role):
        self.update_employee_list(role)

    def update_employee_list(self, role):
        # Clear existing items in the list
        for i in reversed(range(self.employee_list_layout.count())):
            self.employee_list_layout.itemAt(i).widget().setParent(None)

        # Filter employees by role (primary or additional)
        filtered_employees = []
        for emp in self.employees:
            if role == "All" or role == emp.employee_type or role in getattr(emp, 'additional_roles', []):
                filtered_employees.append(emp.name)

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
            
            # Role section with add/remove functionality
            role_section = QWidget()
            role_layout = QVBoxLayout(role_section)
            role_layout.setContentsMargins(0, 0, 0, 0)
            
            # Primary role (Role1)
            primary_role_layout = QHBoxLayout()
            primary_role_label = QLabel("Role1 (Primary):")
            primary_role_input = QComboBox()
            primary_role_input.addItems(list(ROLE_RULES.keys()))
            primary_role_input.setCurrentText(old_employee.employee_type)
            primary_role_layout.addWidget(primary_role_label)
            primary_role_layout.addWidget(primary_role_input)
            
            # Add button for additional roles
            add_role_btn = QPushButton("+")
            add_role_btn.setFixedWidth(30)
            primary_role_layout.addWidget(add_role_btn)
            
            role_layout.addLayout(primary_role_layout)
            
            # Container for additional roles
            additional_roles_container = QWidget()
            additional_roles_layout = QVBoxLayout(additional_roles_container)
            additional_roles_layout.setContentsMargins(0, 0, 0, 0)
            role_layout.addWidget(additional_roles_container)
            
            # List to track additional role inputs
            additional_role_inputs = []
            
            # Function to add a new role input
            def add_role_input(default_role=None):
                role_row = QHBoxLayout()
                role_label = QLabel(f"Role{len(additional_role_inputs) + 2}:")
                role_combo = QComboBox()
                role_combo.addItems(list(ROLE_RULES.keys()))
                if default_role:
                    role_combo.setCurrentText(default_role)
                
                remove_btn = QPushButton("-")
                remove_btn.setFixedWidth(30)
                
                role_row.addWidget(role_label)
                role_row.addWidget(role_combo)
                role_row.addWidget(remove_btn)
                
                # Create a container for this row
                row_container = QWidget()
                row_container.setLayout(role_row)
                
                additional_roles_layout.addWidget(row_container)
                additional_role_inputs.append((row_container, role_combo))
                
                # Connect remove button
                remove_btn.clicked.connect(lambda: remove_role_input(row_container))
            
            # Function to remove a role input
            def remove_role_input(container):
                # Find and remove from tracking list
                for i, (cont, combo) in enumerate(additional_role_inputs):
                    if cont == container:
                        additional_role_inputs.pop(i)
                        break
                
                # Remove from UI
                container.setParent(None)
                
                # Update labels for remaining roles
                for i, (cont, combo) in enumerate(additional_role_inputs):
                    cont.layout().itemAt(0).widget().setText(f"Role{i + 2}:")
            
            # Add existing additional roles
            for role in getattr(old_employee, 'additional_roles', []):
                add_role_input(role)
            
            # Connect add button
            add_role_btn.clicked.connect(lambda: add_role_input())
            
            # Time inputs
            start_time_input = QLineEdit(old_employee.start_time or "")
            end_time_input = QLineEdit(old_employee.end_time or "")
            
            # Build the form
            layout.addWidget(QLabel("Name:"))
            layout.addWidget(name_input)
            layout.addWidget(role_section)
            layout.addWidget(QLabel("Start Time (HH:MM):"))
            layout.addWidget(start_time_input)
            layout.addWidget(QLabel("End Time (HH:MM):"))
            layout.addWidget(end_time_input)
            
            # Add role change listener for primary role
            def primary_role_changed(role):
                is_freelancer = (role == "Freelancer")
                start_time_input.setEnabled(not is_freelancer)
                end_time_input.setEnabled(not is_freelancer)
            
            primary_role_input.currentTextChanged.connect(primary_role_changed)
            primary_role_changed(primary_role_input.currentText())  # Initial check
            
            # Save button
            save_btn = QPushButton("Save")
            save_btn.clicked.connect(lambda: self.save_edited_employee(
                dialog, 
                old_employee.name, 
                name_input.text(), 
                primary_role_input.currentText(),
                [combo.currentText() for _, combo in additional_role_inputs],
                start_time_input.text(), 
                end_time_input.text()
            ))
            
            layout.addWidget(save_btn)
            dialog.exec_()


    def save_edited_employee(self, dialog, old_name, new_name, new_role, additional_roles=None, new_start_time=None, new_end_time=None):
        if new_role != 'Freelancer' and (not new_start_time or not new_end_time):
            QMessageBox.warning(self, "Error", "Fulltimers must provide start and end times")
            return
        
        edit_employee(old_name, new_name, new_role, additional_roles, new_start_time, new_end_time)
        self.employees = load_employees()
        self.availability = load_data()
        self.update_employee_list(self.role_combo.currentText())
        self.update_calendar()
        dialog.accept()

    def save_data(self):
        sync_availability()  # Force synchronization
        save_data(self.availability)
        QMessageBox.information(self, "Saved", "Availability saved!")
    
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
            
            # Get the schedule data
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

    def import_from_excel(self, file_path=None):
        try:
            from PySide6.QtWidgets import QFileDialog
            
            # Open file dialog if no path is provided
            if not file_path:
                file_path, _ = QFileDialog.getOpenFileName(
                    self,
                    "Select Availability Excel File",
                    "",
                    "Excel Files (*.xlsx *.xls)"
                )
            
            # Only proceed if user selected a file
            if file_path:
                from scheduling_logic import import_from_excel
                result = import_from_excel(file_path)
                self.availability = load_data()  # Reload the data after import
                self.update_calendar()  # Update the UI with the new data
                QMessageBox.information(self, "Success", result)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to import data: {str(e)}")


    def export_availability_to_excel(self):
        try:
            from PySide6.QtWidgets import QFileDialog
            
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Export Availability to Excel",
                "availability_export.xlsx",  # Default filename
                "Excel Files (*.xlsx)"
            )
            
            if file_path:  # Only proceed if user didn't cancel
                # Add .xlsx extension if not already present
                if not file_path.endswith('.xlsx'):
                    file_path += '.xlsx'
                    
                result = export_availability_to_excel(self.availability, file_path)
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
    
    def create_debug_package(self):
        try:
            log_info("Creating data package requested by user")
            package_path = create_data_package()
            
            if package_path:
                QMessageBox.information(
                    self, 
                    "Data Package Created", 
                    f"Data package has been created at: {package_path}\n\n"
                    "Please send this file to the developer for debugging."
                )
            else:
                QMessageBox.warning(
                    self,
                    "Error",
                    "Failed to create data package. Please check the logs."
                )
        except Exception as e:
            log_error("Failed to create data package", e)
            QMessageBox.critical(
                self,
                "Error",
                f"An unexpected error occurred: {str(e)}"
            )


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
        export_btn = QPushButton("Export")
        export_btn.clicked.connect(self.export_to_excel)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        
        button_box.addButton(export_btn, QDialogButtonBox.AcceptRole)
        button_box.addButton(cancel_btn, QDialogButtonBox.RejectRole)
        
        layout.addWidget(button_box)
    
    def export_to_excel(self):
        try:
            from PySide6.QtWidgets import QFileDialog
            
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Export Schedule to Excel",
                "schedule.xlsx",  # Default filename
                "Excel Files (*.xlsx)"
            )
            
            if file_path:  # Only proceed if user didn't cancel
                # Add .xlsx extension if not already present
                if not file_path.endswith('.xlsx'):
                    file_path += '.xlsx'
                    
                df = pd.DataFrame(self.schedule_data)
                df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Success", f"Excel schedule has been successfully generated at {file_path}!")
                self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export schedule: {str(e)}")




class LeaveDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Leave Type")
        layout = QVBoxLayout(self)
        
        self.leave_type_combo = QComboBox()
        self.leave_type_combo.addItems(["AL", "CL", "PH", "ON", "自由調配", "half off"])
        layout.addWidget(self.leave_type_combo)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_leave_type(self):
        return self.leave_type_combo.currentText()
    


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        
        # Create and install translator
        translator = QTranslator()
        if not translator.load("zh_TW.ts"):
            if not translator.load("zh_TW.qm"):
                print("Failed to load translation file")
        app.installTranslator(translator)
        
        # Initialize logging before creating the main window
        from logger_utils import setup_logging, log_info, log_error
        log_file = setup_logging()
        log_info("Application started")
        
        window = AvailabilityEditor()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        # This will catch exceptions during startup
        from logger_utils import log_error
        log_error("Fatal error during application startup", e)
        
        # Show error message to user
        if 'app' in locals():
            QMessageBox.critical(None, "Fatal Error", 
                                f"A fatal error occurred during startup: {str(e)}\n\n"
                                f"Please check the log file at: {log_file}")



