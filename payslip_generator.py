import pandas as pd
try:
    df= pd.read_excel("employees.xlsx")
    df["Net Salary"] = df["BASIC SALARY"]+ df["ALLOWANCES"]- df["DEDUCTIONS"]
    print("Employee data loaded successfully")
    print(df.head())

except FileNotFoundError:
    print("Error:'employees.xlsx' file not found")
except Exception as e:
    print(f"An error occured: {e}")



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
    pdf.cell(100, 10, f"${employee['BASIC SALARY']:.2f}", 0, ln=True)

    pdf.cell(50, 10, f"ALLOWANCES:", 0)
    pdf.cell(100, 10, f"${employee['ALLOWANCES']:.2f}", 0, ln=True)

    pdf.cell(50, 10, f"DEDUCTIONS:", 0)
    pdf.cell(100, 10, f"${employee['DEDUCTIONS']:.2f}", 0, ln=True)

    pdf.cell(50, 10, f"Net Salary:", 0)
    pdf.cell(100, 10, f"${employee['Net Salary']:.2f}", 0, ln=True)

    # Save the file
    filename = f"payslips/{employee['EMPLOYEE ID']}.pdf"
    pdf.output(filename)
    print(f"Payslip generated for {employee['NAME']} → {filename}")
for _, employee in df.iterrows():
    generate_payslip(employee)

import yagmail
import os
email_user = os.getenv("EMAIL_USER")
email_password = os.getenv("EMAIL_PASSWORD")
yag = yagmail.SMTP(user=email_user, password=email_password)
def send_payslip(employee):
    subject = "Your Payslip For This Month"
    body =f"Dear{employee['NAME']},\n\nPlease find your attached payslip for this month.\n\nBest regards,\nWaMambo Holdings"
    attachment = f"payslips/{employee['EMPLOYEE ID']}.pdf"

try:
        yag.send(to=employee["EMAIL"], subject=subject, contents=body, attachments=attachment)
        print(f"Payslip sent to {employee['Name']} at {employee['Email']}")
except Exception as e:
        print(f"Error sending email to {employee['Name']}: {e}")

# Send payslips to all employees
for _, employee in df.iterrows():
    send_payslip(employee)
    