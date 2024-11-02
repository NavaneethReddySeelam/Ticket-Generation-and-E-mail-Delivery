import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
import re
import json
from fpdf import FPDF
from dotenv import load_dotenv
import logging

# Load environment variables from .env file
load_dotenv()

# Set up logging
logging.basicConfig(level=logging.INFO)

# Set email credentials
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
SMTP_SERVER = 'smtp.office365.com'
SMTP_PORT = 587

# Function to validate email format
def is_valid_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email) is not None

# Function to create a PDF ticket
def create_ticket_pdf(name, token, date, folder_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_fill_color(0, 0, 0)  # Black background
    pdf.rect(0, 0, 210, 297, 'F')  # A4 size

    # Title and styles
    pdf.set_font("Arial", 'B', 24)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(0, 20, "GamingCraft Hackathon 2024", ln=True, align='C')

    # Content
    pdf.set_font("Arial", 'B', 16)
    pdf.set_text_color(0, 255, 255)
    pdf.cell(0, 20, f"Name: {name}", ln=True, align='C')

    pdf.set_font("Arial", 'B', 18)
    pdf.set_text_color(255, 102, 255)
    pdf.cell(0, 20, f"Token Number: {token}", ln=True, align='C')
    pdf.set_text_color(255, 204, 0)
    pdf.cell(0, 20, f"Date: {date}", ln=True, align='C')

    pdf.set_font("Arial", 'I', 12)
    pdf.set_text_color(200, 200, 200)
    pdf.cell(0, 20, "Join us for an exciting hackathon!", ln=True, align='C')

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    pdf_file_name = os.path.join(folder_path, f"{name.replace(' ', '_')}_ticket.pdf")
    pdf.output(pdf_file_name)
    return pdf_file_name

# Load your Excel file
df = pd.read_excel('participants.xlsx')

# Initialize token
starting_token = 1000

def send_email(name, email, token, pdf_file_name):
    message = MIMEMultipart()
    message['From'] = EMAIL_ADDRESS
    message['To'] = email
    message['Subject'] = 'Your Hackathon Participation Token'
    
    group_link = "https://linktr.ee/klef.mayavi"
    body = f"""Hi {name},
    
Thank you for registering for our Hackathon!

Your unique participation token is: {token}
Please find your ticket attached.

Join our group for updates: {group_link}

Best regards,
Team Mayavi
"""
    message.attach(MIMEText(body, 'plain'))

    with open(pdf_file_name, 'rb') as pdf_file:
        part = MIMEApplication(pdf_file.read(), Name=os.path.basename(pdf_file_name))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(pdf_file_name)}"'
        message.attach(part)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, email, message.as_string())
            logging.info(f"Email sent to {email}")
            return True
    except Exception as e:
        logging.error(f"Failed to send email to {email}: {e}")
        return False

failed_emails = []
hackathon_date = "8-11-2024 to 9-11-2024"
pdf_folder_path = 'tickets'

for index, row in df.iterrows():
    token = starting_token + index
    df.at[index, 'Token'] = token

    if is_valid_email(row['email']):
        pdf_file_name = create_ticket_pdf(row['name'], token, hackathon_date, pdf_folder_path)
        if not send_email(row['name'], row['email'], token, pdf_file_name):
            df.at[index, 'EmailSent'] = 'Failed'
            failed_emails.append(row['name'])
        else:
            df.at[index, 'EmailSent'] = 'Yes'
    else:
        logging.warning(f"Invalid email format for {row['name']}: {row['email']}")

df.to_excel('participants_with_tokens.xlsx', index=False)

if failed_emails:
    failed_df = df[df['EmailSent'] == 'Failed']
    failed_df.to_excel('failed_email_participants.xlsx', index=False)
    logging.info("Failed emails saved.")

json_data = df.to_dict(orient='records')
with open('participants.json', 'w') as json_file:
    json.dump(json_data, json_file, indent=4)
logging.info("JSON file updated with participant data.")
