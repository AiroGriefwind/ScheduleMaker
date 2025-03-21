import json
import pandas as pd
#from collections import deque  
from datetime import datetime, timedelta

class Employee:
    def __init__(self, name, employee_type):
        self.name = name
        self.employee_type = employee_type
        
    def get_available_shifts(self):
        # Get shifts based on employee type from the rule storage
        if self.employee_type in ROLE_RULES:
            rule = ROLE_RULES[self.employee_type]
            if rule["rule_type"] == "shift_based":
                # For shift-based employees like freelancers
                weekday_shifts = list(rule["shifts"]["weekday"].values())
                weekend_shifts = list(rule["shifts"]["weekend"].values())
                return list(set(weekday_shifts + weekend_shifts))
            elif rule["rule_type"] == "fixed_time":
                # For fixed-time employees like senior editors
                shifts = [rule["default_shift"]]
                if "special_duty" in rule:
                    shifts.append(rule["special_duty"]["shift"])
                return shifts
        return []  # Default if no rules found
        
    def get_display_name(self):
        return f"{self.name}({self.employee_type[0]})"


class Freelancer(Employee):
    def __init__(self, name):
        super().__init__(name, "Freelancer")
        
    def get_available_shifts(self):
        return ["7-16", "10-19", "15-24"]

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

def init_availability(start_date, employees):
    return {
        (start_date + timedelta(days=i)).strftime("%Y-%m-%d"): {
            employee.get_display_name(): [] for employee in employees
        } for i in range(7)
    }

def save_data(data):
    with open('availability.json', 'w') as f:
        json.dump(data, f, indent=4, separators=(',', ': '))  # Enhanced readability format


def load_data():
    try:
        with open('availability.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return None

# Constants
EMPLOYEES = init_employees()
FREELANCERS = [employee.get_display_name() for employee in EMPLOYEES if isinstance(employee, Freelancer)]
SHIFT_COLORS = {
    "7-16": (144, 238, 144),   # Light green
    "10-19": (255, 228, 181),  # Light orange
    "15-24": (176, 224, 230)   # Light blue
}
# New centralized role-based rules storage
ROLE_RULES = {
    "Freelancer": {
        "rule_type": "shift_based",
        "shifts": {
            "weekday": {"early": "7-16", "day": "10-19", "night": "15-24"},
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
        "special_duty": {
            "frequency": "monthly",
            "day_of_month": 1,  # 1st day of each month
            "shift": "7-16"
        }
    }
    # Other roles can be added here with their specific rules
}

def add_employee(name, role):
    new_employee = Employee(name=name, employee_type=role)
    EMPLOYEES.append(new_employee)
    save_employees()

def edit_employee(old_name, new_name, new_role):
    for employee in EMPLOYEES:
        if employee.name == old_name:
            employee.name = new_name
            employee.employee_type = new_role
            break
    save_employees()

def delete_employee(name):
    global EMPLOYEES
    EMPLOYEES = [employee for employee in EMPLOYEES if employee.name != name]
    save_employees()

def save_employees():
    with open('employees.json', 'w') as f:
        json.dump([{"name": emp.name, "role": emp.employee_type} for emp in EMPLOYEES], f)

def load_employees():
    try:
        with open('employees.json', 'r') as f:
            data = json.load(f)
            return [Employee(emp["name"], emp["role"]) for emp in data]
    except FileNotFoundError:
        return init_employees()



def generate_schedule(availability, start_date, export_to_excel=True):
    """
    Main function to generate schedules for all employee types.
    Delegates schedule generation to specific functions based on employee type.
    """
    warnings = []
    schedule = []
    
    # Generate schedules for each employee type
    warnings.extend(generate_freelancer_schedule(availability, start_date, schedule))
    warnings.extend(generate_senior_editor_schedule(availability, start_date, schedule))
    
    # Export schedule to Excel if requested
    if export_to_excel:
        df = pd.DataFrame(schedule)
        df.to_excel("schedule_with_senior_editors.xlsx", index=False)
    
    return warnings

def generate_freelancer_schedule(availability, start_date, schedule):
    """
    Generates schedules for freelancers based on their rules and availability.
    """
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
            
            for name in FREELANCERS:
                if shift_time in availability[date.strftime("%Y-%m-%d")][name] and assigned_shifts[name] == 'off':
                    available_shifts = len(availability[date.strftime("%Y-%m-%d")][name])
                    past_assignments = shift_counts[name][shift_name]
                    freelancer_weights[name] = (1 / (available_shifts + 1)) + (1 / (past_assignments + 1))
            
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
    senior_editor_special_duty = senior_editor_rules.get("special_duty", None)
    
    senior_editors = [emp.get_display_name() for emp in EMPLOYEES if emp.employee_type == "SeniorEditor"]
    
    for date in dates:
        date_str = date.strftime("%d/%m/%Y")
        senior_editor_shifts = {}
        
        # Assign shifts to senior editors
        for name in senior_editors:
            if senior_editor_special_duty and date.day == senior_editor_special_duty["day_of_month"]:
                senior_editor_shifts[name] = senior_editor_special_duty["shift"]
            else:
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
    
    save_data(availability)
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
    
    # Load existing availability from JSON
    availability = load_data()  # Assumes load_data() reads from 'availability.json'
    
    # Add Daisy as a Senior Editor to the employee list
    daisy = Employee(name="Daisy", employee_type="SeniorEditor")
    EMPLOYEES.append(daisy)
    print("Employees after adding Daisy:")
    for employee in EMPLOYEES:
        print(employee.get_display_name())
    # Update availability for Daisy while preserving existing data
    for date in availability.keys():
        day = datetime.strptime(date, "%Y-%m-%d").day
        if daisy.get_display_name() not in availability[date]:
            availability[date][daisy.get_display_name()] = []  # Ensure Daisy's key exists
        
        if day == 1:  # Special duty day (e.g., 1st of each month)
            availability[date][daisy.get_display_name()] = ["7-16"]
        else:
            availability[date][daisy.get_display_name()] = ["13-22"]
    
    # Save updated availability back to JSON
    save_data(availability)  # Assumes save_data() writes to 'availability.json'
    export_availability_to_excel(availability)
    # Generate schedule based on updated availability
    warnings = generate_schedule(availability, start_date, export_to_excel=True)
    
    # Print warnings if there are any
    if warnings:
        print("Warnings:")
        for warning in warnings:
            print(warning)
    else:
        print("Schedule generated successfully!")


    









