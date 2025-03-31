import json
import pandas as pd
#from collections import deque  
from datetime import datetime, timedelta

class Employee:
    def __init__(self, name, employee_type, start_time=None, end_time=None):
        self.name = name
        self.employee_type = employee_type
        self.start_time = start_time
        self.end_time = end_time

    def get_available_shifts(self):
        if self.employee_type in ROLE_RULES:
            rule = ROLE_RULES[self.employee_type]
            if rule["rule_type"] == "shift_based":
                weekday_shifts = list(rule["shifts"]["weekday"].values())
                weekend_shifts = list(rule["shifts"]["weekend"].values())
                return list(set(weekday_shifts + weekend_shifts))
            elif rule["rule_type"] == "fixed_time":
                 # Return formatted shift time even when empty
                if self.start_time and self.end_time:
                    return [f"{self.start_time}-{self.end_time}"]
                return [rule.get("default_shift", "Shift Not Set")]
        return []

        


class Freelancer(Employee):
    def __init__(self, name):
        super().__init__(name, "Freelancer")
        
    def get_available_shifts(self):
        # Get current day of week to determine if it's a weekday or weekend
        today = datetime.now().weekday()
        day_type = "weekday" if today < 5 else "weekend"
        
        # Return all shifts for the current day type
        return list(ROLE_RULES["Freelancer"]["shifts"][day_type].values()) 

class SeniorEditor(Employee):
    def __init__(self, name):
        super().__init__(name, "SeniorEditor")
    
    def get_available_shifts(self):
        return ["13-22"]
class economics(Employee):
    def __init__(self, name):
        super().__init__(name, "economics")
    
    def get_available_shifts(self):
        return ["10-19"]
class Entertainment(Employee):
    def __init__(self, name):
        super().__init__(name, "Entertainment")
    
    def get_available_shifts(self):
        return ["10-19"]
class KoreanEntertainment(Employee):
    def __init__(self, name):
        super().__init__(name, "KoreanEntertainment")
    def get_available_shifts(self):
        return ["10-19"]

def init_employees():
    try:
        with open('employees.json', 'r') as f:
            data = json.load(f)
            employees = []
            for emp in data:
                if emp['role'] == 'Freelancer':
                    employees.append(Freelancer(emp['name']))  # ← Proper subclass instantiation
                elif emp['role'] == 'SeniorEditor':
                    employees.append(SeniorEditor(emp['name']))  # ← Add similar for other roles
                else:
                    employees.append(Employee(emp['name'], emp['role']))
            return employees
    except FileNotFoundError:
        return []

def init_availability(start_date, employees):
    return {
        (start_date + timedelta(days=i)).strftime("%Y-%m-%d"): {
            employee.name: [] for employee in employees  # Use employee.name
        } for i in range(7)
    }

def save_data(data):
    try:
        with open('availability.json', 'w') as f:
            json.dump(data, f, indent=4)
    except IOError as e:
        raise RuntimeError(f"File write failed: {str(e)}")[1]



def load_data():
    try:
        with open('availability.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return None

# Constants
EMPLOYEES = init_employees()
FREELANCERS = [employee.name for employee in EMPLOYEES if isinstance(employee, Freelancer)]

# New centralized role-based rules storage
ROLE_RULES = {
    "Freelancer": {
        "rule_type": "shift_based",
        "shifts": {
            "weekday": {"early": "7-16", "day": "0930-1830", "night": "15-24"},
            "weekend": {"early": "7-16", "day": "10-19", "night": "15-24"}
        },
        "requirements": {
            "weekday": {"early": 1, "day": 1, "night": 2},
            "weekend": {"early": 1, "day": 1, "night": 1}
        }
    },
    "SeniorEditor": {
        "rule_type": "fixed_time",
        "default_shift": "13-22",
    },
    # Other roles can be added here with their specific rules
    "economics": {
        "rule_type": "fixed_time",
        "default_shift": "10-19",
    },
    "Entertainment": {
        "rule_type": "fixed_time",
        "default_shift": "10-19",
    },
    "KoreanEntertainment": {
        "rule_type": "fixed_time",
        "default_shift": "10-19",
    }
}

def load_employees():
    try:
        with open('employees.json', 'r') as f:
            data = json.load(f)
            return [Employee(emp["name"], emp["role"], emp.get("start_time"), emp.get("end_time")) for emp in data]
    except FileNotFoundError:
        return init_employees()



def sync_availability():
    employees = load_employees()
    availability = load_data()
    
    # Get current employee names
    current_employees = {e.name for e in employees}  # Use e.name
    
    # Update availability for each date
    for date in availability:
        # Remove deleted employees
        for emp_name in list(availability[date].keys()):
            if emp_name not in current_employees:
                del availability[date][emp_name]
                
        # Add new employees
        for emp in employees:
            if emp.name not in availability[date]:  # Use emp.name
                availability[date][emp.name] = []  # Use emp.name
    
    save_data(availability)

def validate_synchronization():
    employees = load_employees()
    availability = load_data()
    
    assert len(employees) > 0, "No employees found"
    
    # Check all employees exist in availability
    for emp in employees:
        display_name = emp.name
        for date in availability:
            assert display_name in availability[date], \
                f"{display_name} missing from {date}"
    
    # Check for orphaned availability entries
    all_displays = {emp.name for emp in employees}
    for date in availability:
        for emp_name in availability[date]:
            assert emp_name in all_displays, \
                f"Orphaned entry: {emp_name} on {date}"


def save_employees():
    with open('employees.json', 'w') as f:
        json.dump([{
            "name": emp.name,
            "role": emp.employee_type,
            "start_time": emp.start_time,
            "end_time": emp.end_time
        } for emp in EMPLOYEES], f, indent=4, separators=(',', ': '))


def add_employee(name, role, start_time=None, end_time=None):
    if role == 'Freelancer':
        new_emp = Freelancer(name)
    else:
        new_emp = Employee(name, role, start_time, end_time)
    
    EMPLOYEES.append(new_emp)
    save_employees()
    
    availability = load_data()
    
    if availability is None:
        availability = init_availability(datetime.now(), [new_emp])
    else:
        for date in availability:
            if role != 'Freelancer' and start_time and end_time:
                availability[date][new_emp.name] = [f"{start_time}-{end_time}"]
            else:
                availability[date][new_emp.name] = []
    save_data(availability)
    
    return new_emp


def edit_employee(old_name, new_name, new_role, new_start_time=None, new_end_time=None):
    for emp in EMPLOYEES:
        if emp.name == old_name:
            emp.name = new_name
            emp.employee_type = new_role
            if new_role != 'Freelancer':
                emp.start_time = new_start_time
                emp.end_time = new_end_time
            break

    availability = load_data()
    for date in availability:
        if old_name in availability[date]:
            current_shifts = availability[date][old_name]
            # Preserve leaves and special codes
            leaves = [s for s in current_shifts if s in {"AL", "CL", "PH", "ON", "自由調配"}]
            
            if new_role != 'Freelancer' and new_start_time and new_end_time:
                new_shift = f"{new_start_time}-{new_end_time}"
                # Only update non-leave days
                availability[date][new_name] = leaves if leaves else [new_shift]
    
    save_data(availability)
    save_employees()
    sync_availability()


def delete_employee(name):
    global EMPLOYEES
    # Find the employee to delete
    employee_to_delete = next((emp for emp in EMPLOYEES if emp.name == name), None)

    if employee_to_delete:
        # Remove employee from EMPLOYEES list
        EMPLOYEES.remove(employee_to_delete)

        # Save the updated employee list to JSON
        save_employees()

        # Update availability data
        availability = load_data()
        for date in availability:
            if employee_to_delete.name in availability[date]:  # Use employee_to_delete.name
                del availability[date][employee_to_delete.name]  # Use employee_to_delete.name
        save_data(availability)
    else:
        print(f"Employee with name {name} not found.")




# Add this at the module level in scheduling_logic.py
_last_generated_schedule = []

def get_last_generated_schedule():
    global _last_generated_schedule
    return _last_generated_schedule

def generate_schedule(availability, start_date, export_to_excel=True):
    """
    Main function to generate schedules for all employee types.
    Delegates schedule generation to specific functions based on employee type.
    """
    global _last_generated_schedule
    warnings = []
    schedule = []
    
    # Generate schedules for each employee type
    warnings.extend(generate_freelancer_schedule(availability, start_date, schedule))
    warnings.extend(generate_senior_editor_schedule(availability, start_date, schedule))
    #warnings.extend(generate_economics_schedule(availability, start_date, schedule))
    #warnings.extend(generate_entertainment_schedule(availability, start_date, schedule))
    #warnings.extend(generate_korean_entertainment_schedule(availability, start_date, schedule))
    
    # Store the generated schedule
    _last_generated_schedule = schedule
    
    # Export schedule to Excel if requested
    if export_to_excel:
        df = pd.DataFrame(schedule)
        df.to_excel("schedule_with_senior_editors.xlsx", index=False)
    
    return warnings


def generate_freelancer_schedule(availability, start_date, schedule):
    warnings = []
    date_strings = sorted(availability.keys())
    dates = [datetime.strptime(d, "%Y-%m-%d") for d in date_strings]
    
    freelancer_rules = ROLE_RULES["Freelancer"]
    shift_counts = {name: {"early": 0, "day": 0, "night": 0} for name in FREELANCERS}
    
    for date in dates:
        day_type = 'weekend' if date.weekday() >= 5 else 'weekday'
        assigned_shifts = {name: 'off' for name in FREELANCERS}
        
        shifts = freelancer_rules["shifts"][day_type]
        shift_requirements = freelancer_rules["requirements"][day_type]
        shifts_by_priority = sorted(
            shift_requirements.items(),
            key=lambda x: x[1],
            reverse=True
        )
        
        for shift_name, required_count in shifts_by_priority:
            shift_time = shifts[shift_name]
            freelancer_weights = {}
            date_str = date.strftime("%Y-%m-%d")
            for name in FREELANCERS:
                if name in availability[date_str]: # Ensure name exists in availability
                    if shift_time in availability[date_str][name] and assigned_shifts[name] == 'off':
                        available_shifts = len(availability[date_str][name])
                        past_assignments = shift_counts[name][shift_name]
                        freelancer_weights[name] = (1 / (available_shifts + 1)) + (1 / (past_assignments + 1))
                else:
                    warnings.append(f"Warning: {name} not found in availability for {date_str}")
            sorted_freelancers = sorted(freelancer_weights.items(), key=lambda x: x[1], reverse=True)
            assigned_count = 0
        
            for name, _ in sorted_freelancers:
                if assigned_count < required_count:
                    assigned_shifts[name] = shift_time
                    shift_counts[name][shift_name] += 1
                    assigned_count += 1
        
            if assigned_count < required_count:
                warnings.append(
                    f"Warning: Shift {shift_name} on {date.strftime('%Y-%m-%d')} is understaffed. "
                    f"Required: {required_count}, Assigned: {assigned_count}."
                )
    
        schedule_entry = {"Date": date.strftime("%d/%m/%Y"), **assigned_shifts}
        schedule.append(schedule_entry)
    
    return warnings


def generate_senior_editor_schedule(availability, start_date, schedule):
    """
    Generates schedules for senior editors based on their rules and availability.
    Integrates senior editor shifts into existing schedule entries.
    """
    warnings = []
    date_strings = sorted(availability.keys())
    dates = [datetime.strptime(d, "%Y-%m-%d") for d in date_strings]
    
    senior_editor_rules = ROLE_RULES["SeniorEditor"]
    senior_editor_shift = senior_editor_rules["default_shift"]
    
    senior_editors = [emp.name for emp in EMPLOYEES if emp.employee_type == "SeniorEditor"]
    
    for date in dates:
        date_str = date.strftime("%d/%m/%Y")
        senior_editor_shifts = {}
        
        # Assign shifts to senior editors
        for name in senior_editors:
            senior_editor_shifts[name] = senior_editor_shift
        
        # Find the existing entry in the schedule for this date
        existing_entry = next((entry for entry in schedule if entry["Date"] == date_str), None)
        
        if existing_entry:
            # Update the existing entry with senior editor shifts
            existing_entry.update(senior_editor_shifts)
        else:
            # Create a new entry if no existing entry is found (should not happen)
            schedule_entry = {"Date": date_str, **senior_editor_shifts}
            schedule.append(schedule_entry)
    
    return warnings


def import_from_google_form(file_path):
    """
    Import employee availability data from Google Form responses Excel file.
    Handles both full-time and freelancer data formats.
    """
    try:
        df = pd.read_excel(file_path)
        
        # Initialize availability data structure if not exists
        availability = load_data() or {}
        
        # Process each response row
        for _, row in df.iterrows():
            # Skip rows without name
            if pd.isna(row.get('名字')):
                continue
                
            employee_name = row['名字']
            employee_type = row.get('請問您是全職還是兼職？')
            
            # Skip if employee type is not specified
            if pd.isna(employee_type):
                continue
                
            # Process columns based on employee type
            if employee_type == '全職':
                # Process full-time employee leave options
                process_fulltime_availability(availability, row, employee_name)
            elif employee_type == '兼職':
                # Process freelancer shift selections
                process_freelancer_availability(availability, row, employee_name)
                
        # Save updated availability data
        save_data(availability)
        return "Google Form data imported successfully!"
    except Exception as e:
        raise ValueError(f"Failed to import Google Form data: {str(e)}")

def process_fulltime_availability(availability, row, employee_name):
    fulltime_cols = [col for col in row.index if col.startswith('全職 [') and ']' in col]
    
    for col in fulltime_cols:
        date_str = col.split('[')[1].split(']')[0]
        date_parts = date_str.split('/')
        if len(date_parts) == 3:
            iso_date = f"{date_parts[2]}-{date_parts[1]}-{date_parts[0]}"
            
            if iso_date not in availability:
                availability[iso_date] = {}
            
            if employee_name not in availability[iso_date]:
                availability[iso_date][employee_name] = []
            
            leave_value = row[col]
            
            if pd.notna(leave_value) and leave_value:
                availability[iso_date][employee_name] = [leave_value]
            else:
                employee = next((emp for emp in EMPLOYEES if emp.name == employee_name), None)
                if employee and employee.employee_type in ROLE_RULES:
                    # Get actual shift from the form response
                    shift_value = row[col]  # New line to capture actual shift
                    if pd.notna(shift_value) and "-" in shift_value:
                        # Update employee configuration
                        start_time, end_time = shift_value.split('-')
                        employee.start_time = start_time
                        employee.end_time = end_time
                        availability[iso_date][employee_name] = [shift_value]
                    else:
                        # Fallback to default shift
                        rule = ROLE_RULES[employee.employee_type]
                        availability[iso_date][employee_name] = [rule["default_shift"]]



def process_freelancer_availability(availability, row, employee_name):
    """Process freelancer availability from form response."""
    # Identify freelancer date columns (format: '兼職 [DD/MM/YYYY]')
    freelancer_cols = [col for col in row.index if col.startswith('兼職 [') and ']' in col]
    
    for col in freelancer_cols:
        # Extract date from column name
        date_str = col.split('[')[1].split(']')[0]
        
        # Convert date format from DD/MM/YYYY to YYYY-MM-DD for internal storage
        date_parts = date_str.split('/')
        if len(date_parts) == 3:
            iso_date = f"{date_parts[2]}-{date_parts[1]}-{date_parts[0]}"
            
            # Create a datetime object to check if it's a weekday or weekend
            date_obj = datetime.strptime(iso_date, "%Y-%m-%d")
            is_weekend = date_obj.weekday() >= 5  # 5 and 6 are Saturday and Sunday
            
            # Initialize date in availability if not exists
            if iso_date not in availability:
                availability[iso_date] = {}
                
            # Initialize employee in date if not exists
            if employee_name not in availability[iso_date]:
                availability[iso_date][employee_name] = []
            
            # Get shift selections
            shift_value = row[col]
            
            # Skip if no shifts selected
            if pd.isna(shift_value):
                continue
                
            # Process shift selections
            if shift_value == '全選':
                # All shifts selected
                availability[iso_date][employee_name] = ["7-16", "0930-1830" if not is_weekend else "10-19", "15-24"]
            else:
                # Parse individual shift selections
                shifts = []
                if '早更' in str(shift_value):
                    shifts.append("7-16")
                if '日更' in str(shift_value):
                    shifts.append("0930-1830" if not is_weekend else "10-17")
                if '夜更' in str(shift_value):
                    shifts.append("15-24")
                    
                availability[iso_date][employee_name] = shifts



def import_from_excel(file_path):
    df = pd.read_excel(file_path)
    required_columns = {'Date', 'Employee', 'Shift'}
    if not required_columns.issubset(df.columns):
        raise ValueError(f"Excel file must contain columns: {required_columns}")
    
    availability = {}
    for _, row in df.iterrows():
        date_str = row['Date']
        employee_name = row['Employee']
        shift = row['Shift']
        
        if date_str not in availability:
            availability[date_str] = {name: [] for name in FREELANCERS}
        
        if employee_name not in availability[date_str]:
            availability[date_str][employee_name] = []
        
        if shift not in availability[date_str][employee_name]:
            availability[date_str][employee_name].append(shift)
        
        # Update employee configuration
        employee = next((emp for emp in EMPLOYEES if emp.name == employee_name), None)
        if employee and employee.employee_type != 'Freelancer':
            if '-' in shift and shift not in ["AL", "CL", "PH", "ON", "自由調配"]:
                start_time, end_time = shift.split('-')
                employee.start_time = start_time
                employee.end_time = end_time
    
    save_data(availability)
    save_employees()  # Save updated employee configurations
    return "Data imported successfully!"

def export_availability_to_excel(availability):
    data = []
    for date, employees in availability.items():
        for employee_name, shifts in employees.items():
            for shift in shifts:
                data.append({"Date": date, "Employee": employee_name, "Shift": shift})
    
    df = pd.DataFrame(data)
    # Set Excel options to prevent excessive whitespace
    df.to_excel("availability_export.xlsx", index=False, engine='openpyxl')
    return "Availability exported successfully to Excel!"


def clear_availability(start_date, employees):
    return init_availability(start_date, employees)



# Main logic for testing
if __name__ == "__main__":
    # Define start date for testing
    start_date = datetime.strptime("2025-03-19", "%Y-%m-%d")
    
    
    
