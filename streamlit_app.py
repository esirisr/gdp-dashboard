from flask import Flask, render_template_string, request, redirect, url_for
import openpyxl
import os
from threading import Thread

# Create the Flask app
app = Flask(__name__)

# Path to the Excel file where data will be saved
EXCEL_FILE = r"C:\Users\admin\Desktop\Path\collected_data.xlsx"

# Check if the Excel file exists; if not, create it with headers
if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Responses"
    headers = [
        "Full_Name", "Location", "Permit_Numbe", "Permit_Class",
        "Issue_Date", "Renewal_Date", "Expire_Date", "Contacts"
    ]
    ws.append(headers)
    wb.save(EXCEL_FILE)

# HTML template with a dropdown menu for Permit_Class
HTML_FORM = """
<!DOCTYPE html>
<html>
<head>
    <title>Data Collection Form</title>
    <style>
        body {
            font-family: 'Poppins', sans-serif;
            background: linear-gradient(135deg, #6a11cb 0%, #2575fc 100%);
            margin: 0;
            padding: 0;
            color: #333;
        }

        .container {
            max-width: 800px;
            margin: 50px auto;
            background: #ffffff;
            padding: 30px 40px;
            border-radius: 20px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.15);
            animation: fadeIn 1s ease-in-out;
        }

        h1 {
            text-align: center;
            color: #4a4a4a;
            margin-bottom: 30px;
            font-weight: 600;
        }

        .form-group {
            display: flex;
            flex-wrap: wrap;
            justify-content: space-between;
            margin-bottom: 20px;
        }

        .form-group label {
            font-weight: bold;
            color: #555;
            width: 100%;
            margin-bottom: 5px;
        }

        .form-group input,
        .form-group select {
            width: 48%;
            padding: 12px 15px;
            margin: 5px 0;
            border: 1px solid #ddd;
            border-radius: 10px;
            box-sizing: border-box;
            font-size: 16px;
            background-color: #f9f9f9;
            transition: all 0.3s ease;
        }

        .form-group input:focus,
        .form-group select:focus {
            border-color: #6a11cb;
            outline: none;
            box-shadow: 0 0 5px rgba(106, 17, 203, 0.5);
            background-color: #fff;
        }

        button {
            display: block;
            width: 100%;
            padding: 15px;
            background-color: #6a11cb;
            color: white;
            font-size: 18px;
            font-weight: 600;
            border: none;
            border-radius: 10px;
            cursor: pointer;
            transition: all 0.3s ease;
            text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.2);
        }

        button:hover {
            background-color: #2575fc;
            box-shadow: 0 5px 15px rgba(37, 117, 252, 0.3);
        }

        .footer {
            text-align: center;
            margin-top: 30px;
            color: #777;
            font-size: 14px;
        }

        .footer a {
            color: #6a11cb;
            text-decoration: none;
            font-weight: 600;
        }

        .footer a:hover {
            text-decoration: underline;
        }

        @keyframes fadeIn {
            from {
                opacity: 0;
                transform: translateY(-20px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        @media (max-width: 600px) {
            .form-group input,
            .form-group select {
                width: 100%;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Data Collection Form</h1>
        <form method="POST">
            <div class="form-group">
                <input type="text" id="full_name" name="full_name" placeholder="Full Name" required>
                <input type="text" id="location" name="location" placeholder="Location" required>
            </div>
            <div class="form-group">
                <input type="text" id="permit_number" name="permit_number" placeholder="Permit Number" required>
                <select id="permit_class" name="permit_class" required>
                    <option value="">Select Permit Class</option>
                    <option value="Khabiir Sare">Khabiir Sare</option>
                    <option value="Khabiir Koronto">Khabiir Koronto</option>
                    <option value="Xirfadle Sare">Xirfadle Sare</option>
                    <option value="Xirfadle Dhexe">Xirfadle Dhexe</option>
                    <option value="Xirfadle Hoose">Xirfadle Hoose</option>
                </select>
            </div>
            <div class="form-group">
                <input type="date" id="issue_date" name="issue_date" required>
                <input type="date" id="renewal_date" name="renewal_date" required>
            </div>
            <div class="form-group">
                <input type="date" id="expire_date" name="expire_date" required>
                <input type="text" id="contacts" name="contacts" placeholder="Contacts (e.g., 252XXXXXXX)" pattern="252[0-9]{7,}" title="Contacts must start with 252 and include at least 7 digits after." required>
            </div>
            <button type="submit">Submit</button>
        </form>

        <div class="footer">
            <p>By submitting this form, you agree to our <a href="#">Terms and Conditions</a>.</p>
        </div>
    </div>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Collect data from the form
        data = [
            request.form.get("full_name"),
            request.form.get("location"),
            request.form.get("permit_number"),
            request.form.get("permit_class"),
            request.form.get("issue_date"),
            request.form.get("renewal_date"),
            request.form.get("expire_date"),
            request.form.get("contacts"),
        ]

        # Append data to the Excel file
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb.active
        ws.append(data)
        wb.save(EXCEL_FILE)

        return redirect(url_for("thank_you"))

    return render_template_string(HTML_FORM)

@app.route("/thank-you")
def thank_you():
    return "<h1>Thank you for your response!</h1>"

# Function to run Flask in a thread
def run_flask():
    app.run()

# Run the Flask app in a separate thread
flask_thread = Thread(target=run_flask)
flask_thread.start()
