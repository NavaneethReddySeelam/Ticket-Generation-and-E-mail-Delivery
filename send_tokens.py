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

# Load or create 'failed' and 'passed' lists
def load_failed_passed_lists():
    failed_df = pd.read_excel('failed_email_participants.xlsx') if os.path.exists('failed_email_participants.xlsx') else pd.DataFrame(columns=['Name', 'Email', 'Status'])
    passed_df = pd.read_excel('passed_email_participants.xlsx') if os.path.exists('passed_email_participants.xlsx') else pd.DataFrame(columns=['Name', 'Email'])
    return failed_df, passed_df

# Save DataFrame to JSON file
def save_to_json(filename, data):
    with open(filename, 'w') as json_file:
        json.dump(data, json_file, indent=4)

# Function to send email with attached PDF ticket
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

# Main process
def main():
    # Load participants from main list
    participants_df = pd.read_excel('participants.xlsx')
    
    # Load or create 'failed' and 'passed' lists
    failed_df, passed_df = load_failed_passed_lists()
    
    # Initialize token
    starting_token = len(passed_df) + 1000  # Start token based on already passed emails
    hackathon_date = "8-11-2024 to 9-11-2024"
    pdf_folder_path = 'tickets'
    unique_emails = set(passed_df['Email'].values)  # Set of already passed emails

    # Process each participant
    for index, row in participants_df.iterrows():
        name, email = row['name'], row['email']

        if email in unique_emails:
            logging.info(f"Skipping {email}, already processed.")
            continue
        
        if is_valid_email(email):
            token = starting_token + index  # Unique token for each participant
            pdf_file_name = create_ticket_pdf(name, token, hackathon_date, pdf_folder_path)
            if send_email(name, email, token, pdf_file_name):
                passed_df = pd.concat([passed_df, pd.DataFrame([{'Name': name, 'Email': email}])], ignore_index=True)
                unique_emails.add(email)
                participants_df.at[index, 'EmailSent'] = 'Yes'
            else:
                failed_df = pd.concat([failed_df, pd.DataFrame([{'Name': name, 'Email': email, 'Status': 'Failed'}])], ignore_index=True)
                participants_df.at[index, 'EmailSent'] = 'Failed'
        else:
            logging.warning(f"Invalid email format for {name}: {email}")
            participants_df.at[index, 'EmailSent'] = 'Invalid Email'

    # Save updated lists
    participants_df.to_excel('participants_with_tokens.xlsx', index=False)
    failed_df.to_excel('failed_email_participants.xlsx', index=False)
    passed_df.to_excel('passed_email_participants.xlsx', index=False)
    
    # Update JSON file
    json_data = participants_df.to_dict(orient='records')
    save_to_json('participants.json', json_data)
    logging.info("JSON file updated with participant data.")

if __name__ == "__main__":
    main()
