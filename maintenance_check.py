import pandas as pd
from datetime import datetime
from flask import Flask, request, jsonify

# Initialize Flask app
app = Flask(__name__)

# Function to check if the maintenance date is within the maintenance window
def check_maintenance_window(requested_date, equipment_name, company_name):
    # Load the maintenance schedule (use your actual schedule file)
    EXCEL_FILE = "maintenance_schedule.xlsx"
    df = pd.read_excel(EXCEL_FILE)

    # Normalize input values for comparison
    equipment_name = equipment_name.strip().lower()
    company_name = company_name.strip().lower()

    # Normalize data for comparison
    df["Normalized Equipment"] = df["Maintenance subject"].astype(str).str.strip().str.lower()
    df["Normalized Company"] = df["Company"].astype(str).str.strip().str.lower()

    # Find the corresponding row in the maintenance planner
    equipment_data = df[(df["Normalized Equipment"] == equipment_name) & (df["Normalized Company"] == company_name)]

    if equipment_data.empty:
        return {"status": "error", "message": f"(❌ No maintenance record found for '{equipment_name}' under '{company_name}'.)"}

    # Extract scheduled inspection months from Q1-Q4
    scheduled_months = []
    scheduled_month_names = []
    for quarter in ["Inspection date Q1", "Inspection date Q2", "Inspection date Q3", "Inspection date Q4"]:
        if pd.notna(equipment_data[quarter].values[0]):  # Check if there is a scheduled month
            month_name = str(equipment_data[quarter].values[0]).strip()
            scheduled_month_names.append(month_name)
            scheduled_months.append(month_name)

    # Check if requested month exists in scheduled months
    requested_month = requested_date.strftime('%B')  # Get month name (e.g., March)
    
    if requested_month in scheduled_months:
        return {"status": "Yes", "message": f"(✅ The requested date {requested_date.date()} is within the maintenance window.)"}
    else:
        # Get the first scheduled month for that equipment
        next_due_month = scheduled_month_names[0] if scheduled_month_names else "Unknown"
        return {"status": f"No - Due in {next_due_month}", "message": f"(❌ The requested date {requested_date.date()} is NOT within the maintenance window. Next due: {next_due_month}.)"}

# Function to process the maintenance request
def process_maintenance_request(input_data):
    # Extract the input fields
    schedule_date = input_data.get("Schedule Date")
    equipment_name = input_data.get("Equipment/Part")
    company_name = input_data.get("Company")
    email_subject = input_data.get("Email Subject")
    status = input_data.get("Status")
    attachment = input_data.get("Attachment")
    
    # Parse the schedule date
    try:
        requested_date = datetime.strptime(schedule_date, "%d/%m/%y")
    except ValueError:
        return {"status": "error", "message": "Invalid schedule date format."}

    # Check the maintenance window
    maintenance_check_result = check_maintenance_window(requested_date, equipment_name, company_name)

    # Return the result with the message and status
    return {
        "Message": maintenance_check_result["message"],
        "Status": maintenance_check_result["status"]
    }

# Flask route to check for maintenance and process conversation tracking
@app.route('/check_maintenance', methods=['POST'])
def check_maintenance_route():
    try:
        # Get data from the request (Data from webhook)
        data = request.get_json()
        equipment_name = data.get('equipment_name')
        requested_date = data.get('requested_date')
        company_name = data.get('company_name')

        # Process maintenance request
        maintenance_check_result = process_maintenance_request({
            "Schedule Date": requested_date,
            "Equipment/Part": equipment_name,
            "Company": company_name,
            "Email Subject": f"{equipment_name} request",
            "Status": "First schedule request",
            "Attachment": "No"
        })

        # Return maintenance check result
        return jsonify({
            "status": "Success",
            "message": "Maintenance check completed successfully.",
            "maintenance_check_result": maintenance_check_result
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": f"An error occurred: {str(e)}"
        })

# Only run app in development mode, but in production (Render), Gunicorn will handle it
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=10000)  # Running Flask app locally for testing
