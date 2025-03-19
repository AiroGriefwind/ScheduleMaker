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
        json.dump(data, f)

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

def generate_schedule(availability, start_date, export_to_excel=True):
    date_strings = sorted(availability.keys())
    dates = [datetime.strptime(d, "%Y-%m-%d") for d in date_strings]
    schedule = []
    warnings = []

    # Track shift counts for fairness across freelancers
    shift_counts = {name: {"early": 0, "day": 0, "night": 0} for name in FREELANCERS}

    for date in dates:
        day_type = 'weekend' if date.weekday() >= 5 else 'weekday'
        assigned_shifts = {name: 'off' for name in FREELANCERS}

        # Get the rules for freelancers from ROLE_RULES
        freelancer_rules = ROLE_RULES["Freelancer"]
        shifts = freelancer_rules["shifts"][day_type]
        shift_requirements = freelancer_rules["requirements"][day_type]

        # Sort shifts by staffing requirements (highest first)
        shifts_by_priority = sorted(
            shift_requirements.items(),
            key=lambda x: x[1],
            reverse=True
        )

        for shift_name, required_count in shifts_by_priority:
            shift_time = shifts[shift_name]

            # Calculate weights for freelancers based on availability and past assignments
            freelancer_weights = {}
            for name in FREELANCERS:
                if shift_time in availability[date.strftime("%Y-%m-%d")][name] and assigned_shifts[name] == 'off':
                    available_shifts = len(availability[date.strftime("%Y-%m-%d")][name])
                    past_assignments = shift_counts[name][shift_name]
                    freelancer_weights[name] = (1 / (available_shifts + 1)) + (1 / (past_assignments + 1))

            # Sort freelancers by calculated weights
            sorted_freelancers = sorted(freelancer_weights.items(), key=lambda x: x[1], reverse=True)
            assigned_count = 0

            # Assign shifts to freelancers based on weights
            for name, _ in sorted_freelancers:
                if assigned_count < required_count:
                    assigned_shifts[name] = shift_time
                    shift_counts[name][shift_name] += 1
                    assigned_count += 1

            # Handle understaffing warnings
            if assigned_count < required_count:
                warnings.append(
                    f"Warning: Shift {shift_name} on {date.strftime('%Y-%m-%d')} is understaffed. "
                    f"Required: {required_count}, Assigned: {assigned_count}."
                )

        # Add the day's schedule to the overall schedule
        schedule.append({"Date": date.strftime("%d/%m/%Y"), **assigned_shifts})

    # Export schedule to Excel if requested
    if export_to_excel:
        df = pd.DataFrame(schedule)
        df.to_excel("freelancer_schedule_weighted_randomization.xlsx", index=False)

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

# def validate_schedule(availability):
#     warnings = []
#     dates = sorted(availability.keys())
    
#     for date in dates:
#         day_type = 'weekend' if datetime.strptime(date, "%Y-%m-%d").weekday() >= 5 else 'weekday'
#         shift_requirements = SHIFT_REQUIREMENTS[day_type]
        
#         # Track assigned shifts
#         assigned_shifts = {name: 'off' for name in availability[date].keys()}
        
#         for shift_name, required_count in shift_requirements.items():
#             shift_time = SHIFT_MAPPING[shift_name]
            
#             # Find freelancers available for this shift who haven't been assigned yet
#             available_for_shift = [
#                 name for name, shifts in availability[date].items()
#                 if shift_time in shifts and assigned_shifts[name] == 'off'
#             ]
            
#             # Simulate assigning freelancers to this shift
#             assigned_count = 0
#             while assigned_count < required_count and available_for_shift:
#                 freelancer_name = available_for_shift.pop(0)
#                 assigned_shifts[freelancer_name] = shift_time
#                 assigned_count += 1
            
#             # Check if enough freelancers were assigned
#             if assigned_count < required_count:
#                 warnings.append(
#                     f"Warning: Shift {shift_name} on {date} is understaffed. "
#                     f"Required: {required_count}, Assigned: {assigned_count}."
#                 )
    
#     return warnings



