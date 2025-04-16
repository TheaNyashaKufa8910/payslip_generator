

```markdown
# 🧾 Python Payslip Generator & Email Sender

This Python project automates the generation and emailing of employee payslips based on data provided in an Excel file.

## 📌 Features

- Reads employee details and salary information from `employees.xlsx`
- Calculates net salary:  
  **Net Salary = Basic Salary + Allowances - Deductions**
- Generates individual PDF payslips using **FPDF**
- Sends payslips to employee emails using **yagmail**
- Organizes payslips into a `payslips/` directory

## 🗂️ File Structure

```
📁 project-root/
│
├── employees.xlsx             # Input data file
├── payslips/                  # Generated payslip PDFs
├── web_scraper.py             # (Optional if part of larger project)
├── requirements.txt           # Required libraries
├── README.md                  # This file
└── main.py                    # Script for generating & sending payslips
```

## 📥 Requirements

Make sure you have Python 3.x installed.

Install required packages:

```bash
pip install -r requirements.txt
```

### `requirements.txt` should include:

```txt
pandas
fpdf
yagmail
openpyxl
```

## ⚙️ Setup

1. **Add your email credentials as environment variables:**

   > ⚠️ **Security Note:** Avoid hardcoding your Gmail credentials into the script.

   On Windows PowerShell:
   ```powershell
   $env:EMAIL_USER = "youremail@gmail.com"
   $env:EMAIL_PASS = "your_app_password"
   ```

   In the code, access them like this:
   ```python
   email_user = os.getenv("EMAIL_USER")
   email_password = os.getenv("EMAIL_PASS")
   ```

2. **Prepare your `employees.xlsx` file** with the following columns:

| EMPLOYEE ID | NAME | BASIC SALARY | ALLOWANCES | DEDUCTIONS | EMAIL |
|-------------|------|---------------|------------|------------|-------|

## 🚀 How to Run

```bash
python main.py
```

The script will:

1. Load and process the Excel data.
2. Generate payslips for each employee in PDF format.
3. Send payslips to the respective emails.

## 🛡️ Security Tips

- Use **Gmail App Passwords** instead of your real password.
- Never upload credentials or Excel files containing personal data to public repositories.

## 📌 Bonus Ideas (Optional Enhancements)

- Add logging for sent emails and errors.
- Zip all payslips into one archive for backup.
- Send a summary report to admin.

Made by Ravyn Thorne*

