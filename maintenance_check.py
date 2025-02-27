import os
import pandas as pd
from flask import Flask, request, jsonify
from datetime import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# Initialize Flask app
app = Flask(__name__)

# Define the SCOPES for Google Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Define month mapping for easier comparison
month_mapping = {
    "January": 1, "February": 2, "March": 3, "April": 4,
    "May": 5, "June": 6, "July": 7, "August": 8,
    "September": 9, "October": 10, "November": 11, "December": 12
}

# Google Sheets Authentication - Check if token exists
creds = None
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)

# If creds are not available or invalid, initiate the OAuth 2.0 flow
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        # Set the redirect URI to match the one in the Google Cloud Console
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        
        # Ensure the redirect URI matches what you've set in Google Cloud Console
        flow.redirect_uri = 'http://localhost:10000/oauth2callback'
        
        # Start the local server to handle the authentication
        creds = flow.run_local_server(port=5000)

    # Save the credentials for future use
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

# Now you can interact with the Google Sheets API
service = build('sheets', 'v4', credentials=creds)
spreadsheet_id = '1sfHvjSTmDrvHJnSiD_Jll9em15hTXRCD9gcv-Z4Cmnc'  # Updated Google Sheets ID

# Function to check the maintenance schedule in the planner
def check_maintenance(equipment_name, company_name, requested_date):
    """Check if the requested maintenance date matches the correct quarter in the schedule."""

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

# Function to find existing conversation in Google Sheets
def find_conversation(email, subject):
    """Search for a conversation in Google Sheets using email domain and subject."""
    sheet = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="Sheet1!A:F").execute()
    values = sheet.get('values', [])
    
    for row in values:
        if row[1] == email and row[3] == subject:  # Check email and subject
            return row  # Return the row if found

    return None  # No conversation found

# Function to create a new conversation in Google Sheets
def create_new_conversation(email, subject, status):
    """Add a new conversation to Google Sheets."""
    new_data = [
        [email, subject, "", status, datetime.now().strftime('%d/%m/%Y'), f"{email} {subject}"]
    ]
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range="Sheet1!A:F",
        valueInputOption="RAW",
        body={"values": new_data}
    ).execute()

# Flask route to check for conversation and schedule maintenance
@app.route('/check_conversation', methods=['POST'])
def check_conversation():
    # Get data from the request
    data = request.get_json()
    email = data.get('email')
    subject = data.get('subject')
    equipment_name = data.get('equipment_name')
    requested_date = data.get('requested_date')

    # Check maintenance schedule
    maintenance_check_result = check_maintenance(equipment_name, email.split('@')[1], requested_date)

    # Check if conversation exists in Google Sheets
    existing_conversation = find_conversation(email, subject)
    
    if existing_conversation:
        # If found, continue with the existing conversation
        status = existing_conversation[3]
        return jsonify({
            "status": "Conversation found",
            "message": f"Carry on from previous conversation. Current status: {status}",
            "maintenance_check_result": maintenance_check_result
        })
    else:
        # If no conversation is found, create a new entry in Google Sheets
        create_new_conversation(email, subject, "Initial conversation")
        return jsonify({
            "status": "New conversation",
            "message": "Created new conversation in Google Sheets",
            "maintenance_check_result": maintenance_check_result
        })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)  # Runs on port 10000 for Render
