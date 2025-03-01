import os
import pandas as pd
from flask import Flask, request, jsonify
from datetime import datetime

# Initialize Flask app
app = Flask(__name__)

# Function to load induction records from the "induction_records.xlsx" spreadsheet
def load_induction_data(file_path="induction_records.xlsx"):
    df = pd.read_excel(file_path)
    print("Induction Records Columns:", df.columns)  # Debugging line to print column names
    return df

# Function to check if the engineer is inducted
def check_induction_status(company_name, engineer_name):
    # Load the induction data from the spreadsheet
    df = load_induction_data()

    # Normalize column names for consistency (remove spaces and convert to lowercase)
    df.columns = df.columns.str.strip().str.lower()

    # Debugging: print out the first few rows to check the data
    print("Induction Records Sample:", df.head())

    # Normalize inputs
    company_name = company_name.strip().lower()
    engineer_name = engineer_name.strip().lower()

    # Perform case-insensitive string comparison using .str accessor to avoid ambiguity in Series comparison
    engineer_data = df[
        (df["company name"].str.lower() == company_name) & 
        (df["engineer name"].str.lower() == engineer_name)
    ]

    if engineer_data.empty:
        return {"status": "error", "message": f"Engineer '{engineer_name}' from '{company_name}' not found in induction records."}

    # Get the induction expiry date
    expiry_date = pd.to_datetime(engineer_data["induction expiry date"].values[0])
    today = datetime.today()

    if expiry_date >= today:
        return {"status": "inducted", "message": f"Engineer {engineer_name} is inducted. Induction valid until {expiry_date.date()}."}
    else:
        return {"status": "not inducted", "message": f"Engineer {engineer_name} is not inducted. Induction expired on {expiry_date.date()}."}

# Function to find existing conversation in Excel (conversation_tracker.xlsx)
def find_conversation(email, subject):
    try:
        # Load the conversation tracker Excel file
        df = pd.read_excel('conversation_tracker.xlsx')

        # Check for existing conversation (email and subject match)
        existing_conversation = df[(df['Email ID'] == email) & (df['Subject'] == subject)]

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
        df = df.append(new_data, ignore_index=True)

        # Save the updated DataFrame back to the Excel file
        df.to_excel('conversation_tracker.xlsx', index=False)
    
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
        engineer_name = data.get('engineer_name')  # Add engineer name to check induction
        induction_names = data.get('induction_names')  # Check if induction names are provided

        # Check maintenance schedule
        maintenance_check_result = check_maintenance(equipment_name, company_name, requested_date)

        # Only check induction status if induction names are provided
        induction_status = {"status": "skipped", "message": "Induction names not provided."}
        if induction_names is not None:  # Only check if induction names are provided
            induction_status = check_induction_status(company_name, engineer_name)

        # Simulating email and subject for this test (you can replace these)
        email = "test@example.com"  # Add email from your system
        subject = f"{equipment_name} request"

        # Check if conversation exists in Excel
        existing_conversation = find_conversation(email, subject)

        if existing_conversation:
            # If found, continue with the existing conversation
            status = existing_conversation['Status']
            return jsonify({
                "status": "Conversation found",
                "message": f"Carry on from previous conversation. Current status: {status}",
                "maintenance_check_result": maintenance_check_result,
                "induction_status": induction_status
            })
        else:
            # If no conversation is found, create a new entry in Excel
            create_new_conversation(email, subject, "Initial conversation")
            return jsonify({
                "status": "New conversation",
                "message": "Created new conversation in Excel",
                "maintenance_check_result": maintenance_check_result,
                "induction_status": induction_status
            })
    
    except Exception as e:
        return jsonify({
            "status": "error",
            "message": f"An error occurred: {str(e)}"
        })

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=10000)  # Runs on port 10000 for Render
