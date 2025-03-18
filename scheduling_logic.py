import json
import pandas as pd
#from collections import deque  
from datetime import datetime, timedelta

class Employee:
    def __init__(self, name, employee_type):
        self.name = name
        self.employee_type = employee_type
        
    def get_available_shifts(self):
        pass
    
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
RULES = {
    "weekday": {"early": "7-16", "day": "10-19", "night": "15-24"},
    "weekend": {"early": "7-16", "day": "10-19", "night": "15-24"}
}

SHIFT_REQUIREMENTS = {
        "weekday": {"early": 1, "day": 1, "night": 2},
        "weekend": {"early": 1, "day": 1, "night": 1}
    }
    
SHIFT_MAPPING = {
        "early": "7-16",
        "day": "10-19",
        "night": "15-24"
    }

def generate_schedule(availability, start_date, export_to_excel=True):
    date_strings = sorted(availability.keys())
    dates = [datetime.strptime(d, "%Y-%m-%d") for d in date_strings]
    schedule = []
    warnings = []

    freelancer_names = [e.get_display_name() for e in EMPLOYEES if isinstance(e, Freelancer)]

    for date in dates:
        day_type = 'weekend' if date.weekday() >= 5 else 'weekday'
        assigned_shifts = {name: 'off' for name in FREELANCERS}

        # Sort shifts by staffing requirements (highest first)
        shifts_by_priority = sorted(
            SHIFT_REQUIREMENTS[day_type].items(),
            key=lambda x: x[1],
            reverse=True
        )

        for shift_name, required_count in shifts_by_priority:
            shift_time = SHIFT_MAPPING[shift_name]
            # Find freelancers available for this shift who haven't been assigned yet
            available_for_shift = [
                name for name in FREELANCERS 
                if shift_time in availability[date.strftime("%Y-%m-%d")][name]
                and assigned_shifts[name] == 'off'
            ]

            if len(available_for_shift) < required_count:
                # Not enough freelancers available for this shift
                warnings.append(
                    f"Warning: Shift {shift_name} on {date.strftime('%Y-%m-%d')} is understaffed. "
                    f"Required: {required_count}, Available: {len(available_for_shift)}."
                )
                # Assign whoever is available
                for freelancer_name in available_for_shift:
                    assigned_shifts[freelancer_name] = shift_time
            else:
                # Assign required number of freelancers to this shift
                for i in range(required_count):
                    assigned_shifts[available_for_shift[i]] = shift_time

        schedule.append({"Date": date.strftime("%d/%m/%Y"), **assigned_shifts})

    # Export to Excel if requested
    if export_to_excel:
        df = pd.DataFrame(schedule)
        df.to_excel("freelancer_schedule.xlsx", index=False)

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



