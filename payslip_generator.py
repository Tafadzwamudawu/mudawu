import pandas as pd
from fpdf import FPDF
import yagmail
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

# Directory to save payslips
os.makedirs("payslips", exist_ok=True)

# Read employee data from Excel
try:
    df = pd.read_excel("employees.xlsx")
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit(1)

# Connect to email server
try:
    yag = yagmail.SMTP(EMAIL_USER, EMAIL_PASS)
except Exception as e:
    print(f"Failed to login to email: {e}")
    exit(1)

# Loop through each employee and process their payslip
for index, row in df.iterrows():
    try:
        emp_id = str(row["Employee ID"])
        name = row["Name"]
        email = row["Email"]
        basic = float(row["Basic Salary"])
        allowances = float(row["Allowances"])
        deductions = float(row["Deductions"])
        net_salary = basic + allowances - deductions

        # Create PDF payslip
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        pdf.cell(200, 10, txt="Monthly Payslip", ln=True, align="C")
        pdf.ln(10)
        pdf.cell(100, 10, txt=f"Employee ID: {emp_id}", ln=True)
        pdf.cell(100, 10, txt=f"Name: {name}", ln=True)
        pdf.ln(5)
        pdf.cell(100, 10, txt=f"Basic Salary: ${basic:.2f}", ln=True)
        pdf.cell(100, 10, txt=f"Allowances: ${allowances:.2f}", ln=True)
        pdf.cell(100, 10, txt=f"Deductions: ${deductions:.2f}", ln=True)
        pdf.cell(100, 10, txt=f"Net Salary: ${net_salary:.2f}", ln=True)
        pdf.ln(10)
        pdf.cell(100, 10, txt="Thank you for your hard work!", ln=True)

        # Save PDF
        pdf_path = f"payslips/{emp_id}.pdf"
        pdf.output(pdf_path)

        # Send payslip via email
        subject = "Your Payslip for This Month"
        body = f"Dear {name},\n\nPlease find attached your payslip for this month.\n\nRegards,\nHR Department"
        yag.send(to=email, subject=subject, contents=body, attachments=pdf_path)

        print(f"✅ Payslip sent to {name} ({email})")

    except Exception as e:
        print(f"❌ Error processing employee {row.get('Name', 'Unknown')}: {e}")
