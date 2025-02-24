import os
import pandas as pd
from flask import Flask, request, jsonify
from datetime import datetime

app = Flask(__name__)

# Secure API key - Set this in Render as an environment variable
API_KEY = os.getenv("API_KEY", "your-secret-key")  # Change "your-secret-key" to a strong key

# Load the maintenance planner
EXCEL_FILE = "maintenance_schedule.xlsx"  # Ensure this file is in the correct directory
df = pd.read_excel(EXCEL_FILE)

# Convert months into a dictionary for easy comparison
month_mapping = {
    "January": 1, "February": 2, "March": 3, "April": 4,
    "May": 5, "June": 6, "July": 7, "August": 8,
    "September": 9, "October": 10, "November": 11, "December": 12
}

def check_maintenance(equipment_name, company_name, requested_date):
    """Check if the requested maintenance date matches the correct quarter in the schedule."""
    
    # Convert requested date to a datetime object
    requested_date = pd.to_datetime(requested_date, dayfirst=True)
    requested_month = requested_date.month
    requested_year = requested_date.year

    # Normalize input values for comparison (lowercase + strip spaces)
    equipment_name = equipment_name.strip().lower()
    company_name = company_name.strip().lower()

    # Normalize Excel data for comparison
    df["Normalized Equipment"] = df["Maintenance subject"].astype(str).str.strip().str.lower()
    df["Normalized Company"] = df["Company"].astype(str).str.strip().str.lower()

    # Find the corresponding row in the maintenance planner
    equipment_data = df[
        (df["Normalized Equipment"] == equipment_name) & 
        (df["Normalized Company"] == company_name)
    ]

    if equipment_data.empty:
        return {"status": "error", "message": f"(❌ No maintenance record found for '{equipment_name}' under '{company_name}'.)"}

    # Extract scheduled inspection months from Q1-Q4
    scheduled_months = []
    scheduled_month_names = []
    for quarter in ["Inspection date Q1", "Inspection date Q2", "Inspection date Q3", "Inspection date Q4"]:
        if pd.notna(equipment_data[quarter].values[0]):  # Check if there is a scheduled month
            month_name = str(equipment_data[quarter].values[0]).strip()
            scheduled_month_names.append(month_name)
            scheduled_months.append(month_mapping.get(month_name, None))

    # Check if requested month exists in scheduled months
    if requested_month in scheduled_months:
        return {"status": "Yes", "message": f"(✅ The requested date {requested_date.date()} is within the maintenance window.)"}
    else:
        # Get the first scheduled month for that equipment
        next_due_month = scheduled_month_names[0] if scheduled_month_names else "Unknown"
        return {"status": f"No - Due in {next_due_month}", "message": f"(❌ The requested date {requested_date.date()} is NOT within the maintenance window. Next due: {next_due_month}.)"}

@app.route('/check_maintenance', methods=['POST'])
def check_maintenance_api():
    """API endpoint with authentication."""
    data = request.get_json()
    
    # Get the API key from headers
    api_key = request.headers.get("X-API-Key")

    # Validate API Key
    if api_key != API_KEY:
        return jsonify({"error": "Unauthorized access. Invalid API key."}), 401

    equipment_name = data.get("equipment_name")
    company_name = data.get("company_name")
    requested_date = data.get("requested_date")

    result = check_maintenance(equipment_name, company_name, requested_date)
    return jsonify(result)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)  # Runs on port 10000 for Render
