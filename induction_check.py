import pandas as pd
from flask import Flask, request, jsonify
from datetime import datetime

# Path to the induction tracker Excel file
INDUCTION_FILE = r"C:\Users\stefa\Desktop\AI\maintenance-check\induction_tracker.xlsx"

app = Flask(__name__)

@app.route('/check_inductions', methods=['POST'])
def check_inductions():
    try:
        print("Received request for induction check")

        # Load the induction tracker Excel file
        try:
            df = pd.read_excel(INDUCTION_FILE)
            print("Induction file loaded successfully")
            print("Columns in Excel file:", list(df.columns))
        except Exception as e:
            print(f"Failed to load Excel file: {e}")
            return jsonify({"status": "error", "message": "Could not load induction data."})

        # Parse input data
        data = request.get_json()
        company = data.get("company")
        engineers = data.get("engineers", [])
        maintenance_date_str = data.get("maintenance_date")  # Expecting format 'YYYY-MM-DD'

        print(f"Company: {company}")
        print(f"Engineers: {engineers}")
        print(f"Scheduled Maintenance Date: {maintenance_date_str}")

        if not maintenance_date_str:
            return jsonify({"status": "error", "message": "Missing maintenance date."})

        maintenance_date = datetime.strptime(maintenance_date_str, "%Y-%m-%d").date()

        responses = []

        for engineer in engineers:
            match = df[
                (df['Company'].str.strip().str.lower() == company.strip().lower()) &
                (df['Name'].str.strip().str.lower() == engineer.strip().lower())
            ]

            if match.empty:
                responses.append(f"{engineer} requires an induction.")
            else:
                try:
                    expiry_date = pd.to_datetime(match.iloc[0]['Expiry Date (Auto)']).date()
                    if expiry_date >= maintenance_date:
                        responses.append(f"{engineer} is inducted for the scheduled date.")
                    else:
                        responses.append(f"{engineer}'s induction expires before the scheduled date and needs to be redone.")
                except Exception as e:
                    responses.append(f"Could not parse expiry date for {engineer}.")

        return jsonify({
            "status": "success",
            "results": responses
        })

    except Exception as e:
        print(f"Error in endpoint: {e}")
        return jsonify({"status": "error", "message": f"Server error: {str(e)}"})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=10000)
