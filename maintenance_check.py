import os
import pandas as pd
from flask import Flask, request, jsonify
from datetime import datetime

# Initialize Flask app
app = Flask(__name__)

# Define the month mapping for easier comparison
month_mapping = {
    "January": 1, "February": 2, "March": 3, "April": 4,
    "May": 5, "June": 6, "July": 7, "August": 8,
    "September": 9, "October": 10, "November": 11, "December": 12
}

# Function to check the maintenance schedule in the planner (Excel)
def check_maintenance(equipment_name, company_name, requested_date):
    """Check if the requested maintenance date matches the correct quarter in the schedule."""
    try:
        # Convert requested date to a datetime object
        requested_date = pd.to_datetime(requested_date, dayfirst=True)
        requested_month = requested_date.month
        requested_year = requested_date.year

        # Normalize input values for comparison (lowercase + strip spaces)
        equipment_name = equipment_name.strip().lower()
        company_name = company_name.strip().lower()

        # Load the maintenance planner (Excel file)
        EXCEL_FILE = "maintenance_schedule.xlsx"  # Ensure this file is in the correct directory
        df = pd.read_excel(EXCEL_FILE)

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
    
    except Exception as e:
        return {"status": "error", "message": f"(❌ An error occurred while checking the maintenance: {str(e)})"}

# Function to find existing conversation in Excel (conversation_tracker.xlsx)
def find_conversation(email, subject):
    try:
        # Load the conversation tracker Excel file
        df = pd.read_excel('conversation_tracker.xlsx')

        # Normalize Email and Subject for case-insensitive and stripped comparison
        email = email.strip().lower()
        subject = subject.strip().lower()

        # Check for existing conversation (email and subject match)
        existing_conversation = df[
            (df['Email ID'].str.lower() == email) & (df['Subject'].str.lower() == subject)
        ]

        if not existing_conversation.empty:
            return existing_conversation.iloc[0]  # Return the first matched row
        return None  # No conversation found
    
    except Exception as e:
        return None  # If error occurs, return None

# Function to create a new conversation in Excel (conversation_tracker.xlsx)
def create_new_conversation(email, subject, status):
    try:
        # Load the existing conversation tracker Excel file
        df = pd.read_excel('conversation_tracker.xlsx')

        # Ensure the email is not None or empty
        if not email or pd.isna(email):
            return {"status": "error", "message": "Email address is missing or invalid."}

        # Add a new row with the provided data
        new_data = {
            'Email ID': email,
            'Sender Domain': email.split('@')[1],  # Extract domain from the email
            'Company Name': email.split('@')[1].split('.')[0],  # Basic way to get the company name
            'Subject': subject,
            'Status': status,
            'Last Updated': datetime.now().strftime('%Y-%m-%d'),
            'Sender Domain + Subject': f"{email.split('@')[1]} {subject}"
        }

        # Remove any rows with NaN or empty values in essential columns (Email ID, Subject)
        df.dropna(subset=['Email ID', 'Subject'], inplace=True)

        # Append the new conversation to the DataFrame
        df = df.append(new_data, ignore_index=True)

        # Save the updated DataFrame back to the Excel file
        df.to_excel('conversation_tracker.xlsx', index=False)

        return {"status": "success", "message": "New conversation added successfully!"}
    
    except Exception as e:
        return {"status": "error", "message": f"(❌ Error creating new conversation: {str(e)})"}

# Flask route to check for maintenance and manage conversations
@app.route('/check_maintenance', methods=['POST'])
def check_maintenance_route():
    try:
        # Get data from the request (Data from webhook)
        data = request.get_json()

        equipment_name = data.get('equipment_name')
        requested_date = data.get('requested_date')
        company_name = data.get('company_name')
        email = data.get('email')  # Make sure email is passed in the request
        subject = f"{equipment_name} request"

        # Check maintenance schedule
        maintenance_check_result = check_maintenance(equipment_name, company_name, requested_date)

        # Check if conversation exists in Excel
        existing_conversation = find_conversation(email, subject)

        if existing_conversation:
            # If found, continue with the existing conversation
            status = existing_conversation['Status']
            return jsonify({
                "status": "Conversation found",
                "message": f"Carry on from previous conversation. Current status: {status}",
                "maintenance_check_result": maintenance_check_result
            })
        else:
            # If no conversation is found, create a new entry in Excel
            create_new_conversation(email, subject, "Initial conversation")
            return jsonify({
                "status": "New conversation",
                "message": "Created new conversation in Excel",
                "maintenance_check_result": maintenance_check_result
            })
    
    except Exception as e:
        return jsonify({
            "status": "error",
            "message": f"An error occurred: {str(e)}"
        })

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=10000)  # Runs on port 10000 for Render
