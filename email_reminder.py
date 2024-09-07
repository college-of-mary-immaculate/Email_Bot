import os
import smtplib
import pandas as pd
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path
from dotenv import load_dotenv

PORT = 587
EMAIL_SERVER = "smtp.gmail.com"

current_dir = Path(__file__).resolve().parent if "__file__" in locals() else Path.cwd()
envars = current_dir / ".env"
load_dotenv(envars)

sender_email = os.getenv("EMAIL")
password = os.getenv("PASSWORD")

if not sender_email or not password:
    raise ValueError("Email or password is not set. Check your environment variables.")

def send_email(subject, receiver_email, name, due_date, amount, order):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr(("Yesh Digitals", f"{sender_email}"))
    msg["To"] = receiver_email
    msg["BCC"] = sender_email

    msg.set_content(
        f"""\

        Dear {name},

        I trust this message finds you in good health. As we approach the deadline, I would like to kindly remind you about your recent order, [{order}] Template with Yesh Digitals. The total amount due is {amount} Pesos, and the payment should be settled by {due_date}.
        Your full payment is greatly appreciated and will help ensure a smooth transaction. If you have questions, feel free to message me. Thank you!

        Best regards, 
        Yesh Digitals
        """
    )

    msg.add_alternative(
        f"""\
        <html>
            <body>
                <p>Dear {name}, </p>
                <p>I trust this message finds you in good health. As we approach the deadline, I would like to kindly remind you about your recent order, [{order}] Template with Yesh Digitals. The total amount due is <strong>{amount} Pesos</strong>, and the payment should be settled by <strong>{due_date}</strong>.
                Your full  payment is greatly appreciated and will help ensure a smooth transaction. If you have questions, feel free to message me. Thank you!</p>
                <p>Best regards, </p>
                <p>Yesh Digitals</p>
            </body>
        </html>
        """,
        subtype="html"
    )

    try:
        with smtplib.SMTP(EMAIL_SERVER, PORT) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        print("Email sent successfully to " + receiver_email)
        return True
    except smtplib.SMTPAuthenticationError as e:
        print(f"Authentication failed: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
    return False

if __name__ == "__main__":
    excel_file = "info.xlsx"
    df = pd.read_excel(excel_file)

    if 'Email Sending Date' not in df.columns:
        df['Email Sending Date'] = pd.NaT
    if 'Email Sent' not in df.columns:
        df['Email Sent'] = ''

    current_date = pd.Timestamp(datetime.now().date())  

    for idx, row in df.iterrows():
        due_date = pd.to_datetime(row["Due Date"]).date()
        email_sending_date = due_date - timedelta(days=3)

        
        email_sending_date = pd.to_datetime(email_sending_date)

       
        df.at[idx, 'Email Sending Date'] = email_sending_date

       
        if email_sending_date == current_date and row.get('Email Sent', '') != 'Sent':
            email_sent = send_email(
                subject="Payment Reminder for Recent Order",
                name=row["Name"],
                receiver_email=row["Email"],
                due_date=row["Due Date"],
                amount=row["Amount"],
                order=row["Order"]
            )
            df.at[idx, 'Email Sent'] = 'Sent' if email_sent else 'Failed'
        else:
            if email_sending_date > current_date:
                df.at[idx, 'Email Sent'] = 'Not Yet Sent'
            elif email_sending_date < current_date and row.get('Email Sent', '') != 'Sent':
                df.at[idx, 'Email Sent'] = 'Missed'
    
    
    try:
        df.to_excel(excel_file, index=False)
        print("DataFrame successfully saved to Excel file.")
    except Exception as e:
        print(f"An error occurred while saving the Excel file: {e}")
