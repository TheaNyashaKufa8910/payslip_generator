import pandas as pd
try:
    df = pd.read_excel("employees.xlsx")
    df["BASIC SALARY"] = pd.to_numeric(df["BASIC SALARY"], errors="coerce")
    df["ALLOWANCES"] = pd.to_numeric(df["ALLOWANCES"], errors="coerce")
    df["DEDUCTIONS"] = pd.to_numeric(df["DEDUCTIONS"], errors="coerce")
    df = df.drop(columns=["NET SALARY"], errors="ignore")  # Clean up if exists
    df["NET SALARY"] = df["BASIC SALARY"] + df["ALLOWANCES"] - df["DEDUCTIONS"]

    print("Employee data loaded successfully")
    print(df.head())


except FileNotFoundError:
    print("Error: 'employees.xlsx' file not found")
    df = None
except PermissionError:
    print("Error: Permission denied while accessing 'employees.xlsx'. Close it if it's open in another program.")
    df = None
except Exception as e:
    print(f"An error occurred: {e}")
    df = None




from fpdf import FPDF
import os
os.makedirs("payslips", exist_ok=True)

def generate_payslip(employee):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.set_fill_color(200,220,255)
    pdf.cell(0,10,"Employee Payslip", ln=True, align='C', fill=True)
    pdf.ln(10)

    #employee details
    pdf.cell(50,10, f"EMPLOYEE ID:",0)
    pdf.cell(100,10, str(employee["EMPLOYEE ID"]),0, ln=True)

    pdf.cell(50, 10, f"NAME:",0)
    pdf.cell(100, 10, employee["NAME"],0, ln=True)

    pdf.cell(50, 10, f"BASIC SALARY:", 0)
    pdf.cell(100, 10, f"${float(employee['BASIC SALARY']):,.2f}", 0, ln=True)

    pdf.cell(50, 10, f"ALLOWANCES:", 0)
    pdf.cell(100, 10, f"${float(employee['ALLOWANCES']):,.2f}", 0, ln=True)

    pdf.cell(50, 10, f"DEDUCTIONS:", 0)
    pdf.cell(100, 10, f"${float(employee['DEDUCTIONS']):,.2f}", 0, ln=True)

    pdf.cell(50, 10, f"NET SALARY:", 0)
    pdf.cell(100, 10, f"${float(employee['NET SALARY']):,.2f}", 0, ln=True)

    # Save the file
    filename = f"payslips/{employee['EMPLOYEE ID']}.pdf"
    pdf.output(filename)
    print(f"Payslip generated for {employee['NAME']} → {filename}")
for _, employee in df.iterrows():
    generate_payslip(employee)

import yagmail
import os
email_user = os.getenv("nyashakufa29@gmail.com")
email_password = os.getenv("wijj yygc bmqc vktl")
yag = yagmail.SMTP(user="nyashakufa29@gmail.com", password="wijj yygc bmqc vktl", host='smtp.gmail.com')

def send_payslip(employee):
    subject = "Your Payslip For This Month"
    body = f"Dear  {employee['NAME']},\n\nPlease find your attached payslip for this month.\n\nBest regards,\nWaMambo Holdings"
    attachment = f"payslips/{employee['EMPLOYEE ID']}.pdf"

    try:
        yag.send(to=employee["EMAIL"], subject=subject, contents=body, attachments=attachment)
        print(f"Payslip sent to {employee['NAME']} at {employee['EMAIL']}")
    except Exception as e:
        print(f"Error sending email to {employee['NAME']}: {e}")


# Send payslips to all employees
if df is not None:
    for _, employee in df.iterrows():
        generate_payslip(employee)

    for _, employee in df.iterrows():
        send_payslip(employee)

    #zip_payslips()

    