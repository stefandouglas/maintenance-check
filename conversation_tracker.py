import pandas as pd
from datetime import datetime
from flask import Flask, request, jsonify

CONVERSATION_FILE = r"C:\Users\stefa\Desktop\AI\maintenance-check\conversation_tracker.xlsx"

app = Flask(__name__)

def find_conversation(email, subject):
    try:
        df = pd.read_excel(CONVERSATION_FILE)
        email = email.strip().lower()
        subject = subject.strip().lower()
        existing_conversation = df[
            (df['Email ID'].str.strip().str.lower() == email) &
            (df['Subject'].str.strip().str.lower() == subject)
        ]
        if not existing_conversation.empty:
            return df, existing_conversation.index[0]
        return df, None
    except Exception as e:
        print(f"Error in find_conversation: {str(e)}")
        return None, None

def determine_next_status(current_status, attachment_present, engineer_names_present):
    if current_status.lower() == 'scheduling request':
        return 'Awaiting RAMS and Engineer Names'
    elif current_status.lower() == 'awaiting rams and engineer names':
        if attachment_present and engineer_names_present:
            return 'Conversation Complete'
        elif attachment_present:
            return 'Awaiting Engineer Names'
        elif engineer_names_present:
            return 'Awaiting RAMS'
    elif current_status.lower() == 'awaiting rams' and attachment_present:
        return 'Conversation Complete'
    elif current_status.lower() == 'awaiting engineer names' and engineer_names_present:
        return 'Conversation Complete'
    return current_status

def get_next_step_instruction(status):
    status = status.lower()
    if status == "scheduling request":
        return "Please ask the contractor to provide a proposed maintenance date."
    elif status == "awaiting rams and engineer names":
        return "Please ask the contractor to provide RAMS and the names of the attending engineers."
    elif status == "awaiting engineer names":
        return "Please ask for the names of the attending engineers."
    elif status == "awaiting rams":
        return "Please ask for RAMS."
    elif status == "conversation complete":
        return "All information has been received. Confirm attendance and say thank you."
    else:
        return "Continue monitoring this conversation."

def update_conversation_status(df, index, new_status):
    try:
        df.at[index, 'Status'] = new_status
        df.at[index, 'Last Updated'] = datetime.now().strftime('%Y-%m-%d')
        df.to_excel(CONVERSATION_FILE, index=False)
        return {"status": "success", "message": f"Conversation updated to {new_status}.", "new_status": new_status}
    except Exception as e:
        print(f"Error in update_conversation_status: {str(e)}")
        return {"status": "error", "message": f"Error updating conversation: {str(e)}"}

def determine_initial_status(attachment_present, engineer_names_present):
    if attachment_present and engineer_names_present:
        return "Conversation Complete"
    elif attachment_present:
        return "Awaiting Engineer Names"
    elif engineer_names_present:
        return "Awaiting RAMS"
    else:
        return "Scheduling Request"

def create_new_conversation(email, subject, initial_status):
    try:
        df = pd.read_excel(CONVERSATION_FILE)
        new_data = {
            'Email ID': email,
            'Sender Domain': email.split('@')[1],
            'Company Name': email.split('@')[1].split('.')[0],
            'Subject': subject,
            'Status': initial_status,
            'Last Updated': datetime.now().strftime('%Y-%m-%d'),
            'Sender Domain + Subject': f"{email.split('@')[1]} {subject}"
        }
        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
        df.to_excel(CONVERSATION_FILE, index=False)
        return {"status": "success", "message": "New conversation created.", "new_status": initial_status}
    except Exception as e:
        print(f"Error in create_new_conversation: {str(e)}")
        return {"status": "error", "message": f"Error creating new conversation: {str(e)}"}

@app.route('/check_conversation', methods=['POST'])
def check_conversation():
    try:
        data = request.get_json()
        email = data.get('email')
        subject = data.get('email_subject')
        attachment = data.get('attachment', 'No').strip().lower() == 'yes'
        engineers_raw = data.get('engineer_names', '')
        engineers = [e.strip() for e in engineers_raw.split(',')] if engineers_raw and engineers_raw.lower() != 'none' else []
        engineer_names_present = bool(engineers)

        df, index = find_conversation(email, subject)

        if df is not None and index is not None:
            current_status = df.at[index, 'Status']
            new_status = determine_next_status(current_status, attachment, engineer_names_present)
            instruction = get_next_step_instruction(new_status)
            response_data = update_conversation_status(df, index, new_status)
            response_data["next_step_instruction"] = instruction
            return jsonify(response_data)
        elif df is not None:
            initial_status = determine_initial_status(attachment, engineer_names_present)
            instruction = get_next_step_instruction(initial_status)
            response_data = create_new_conversation(email, subject, initial_status)
            response_data["next_step_instruction"] = instruction
            return jsonify(response_data)
        else:
            return jsonify({"status": "error", "message": "Failed to read conversation tracker."})
    except Exception as e:
        print(f"Error in check_conversation route: {str(e)}")
        return jsonify({"status": "error", "message": f"Error: {str(e)}"})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=10000)
