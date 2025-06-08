import pandas as pd
from datetime import datetime

# Define the conversation tracker file path
CONVERSATION_FILE = r"C:\Users\stefa\Desktop\AI\maintenance-check\conversation_tracker.xlsx"

# Function to check for an existing conversation based on email and subject
def find_conversation(email, subject):
    try:
        # Load the conversation tracker Excel file
        df = pd.read_excel(CONVERSATION_FILE)

        # Normalize email and subject for comparison
        email = email.strip().lower()
        subject = subject.strip().lower()

        # Check for existing conversation (email and subject match)
        existing_conversation = df[(df['Email ID'].str.strip().str.lower() == email) & 
                                   (df['Subject'].str.strip().str.lower() == subject)]

        if not existing_conversation.empty:
            return existing_conversation.iloc[0]  # Return the first matched row
        return None  # No conversation found
    
    except Exception as e:
        return None  # If error occurs, return None

# Function to create a new conversation in the tracker
def create_new_conversation(email, subject, status):
    try:
        # Load the existing conversation tracker Excel file
        df = pd.read_excel(CONVERSATION_FILE)

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
        df.to_excel(CONVERSATION_FILE, index=False)
        return {"status": "success", "message": "New conversation created."}
    
    except Exception as e:
        return {"status": "error", "message": f"Error creating new conversation: {str(e)}"}

# Function to update an existing conversation's status
def update_conversation_status(email, subject, new_status):
    try:
        # Load the existing conversation tracker Excel file
        df = pd.read_excel(CONVERSATION_FILE)

        # Find the row for the existing conversation
        existing_conversation = df[(df['Email ID'].str.strip().str.lower() == email) & 
                                   (df['Subject'].str.strip().str.lower() == subject)]

        if not existing_conversation.empty:
            # Update the status
            df.loc[existing_conversation.index, 'Status'] = new_status
            df.loc[existing_conversation.index, 'Last Updated'] = datetime.now().strftime('%Y-%m-%d')

            # Save the updated DataFrame back to the Excel file
            df.to_excel(CONVERSATION_FILE, index=False)
            return {"status": "success", "message": f"Conversation updated to {new_status}."}
        else:
            return {"status": "error", "message": "Conversation not found."}
    
    except Exception as e:
        return {"status": "error", "message": f"Error updating conversation: {str(e)}"}

# Function to check and update conversations (process incoming data)
def process_conversation_data(email, subject, status):
    existing_conversation = find_conversation(email, subject)

    if existing_conversation is not None:
        # If conversation exists, update the status
        return update_conversation_status(email, subject, status)
    else:
        # If no conversation found, create a new conversation
        return create_new_conversation(email, subject, status)

# Test the system by passing data
input_data = {
    "email": "test@example.com",
    "subject": "Fire alarm maintenance reschedule",
    "status": "Waiting on RAMS"  # Example status to test
}

result = process_conversation_data(input_data["email"], input_data["subject"], input_data["status"])
print(result)
