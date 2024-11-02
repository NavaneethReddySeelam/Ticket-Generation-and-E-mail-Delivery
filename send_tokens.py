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

# Load environment variables from .env file
load_dotenv()  # Load environment variables from .env file

# Set your email credentials
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')  # Ensure you set this in your .env file
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')  # Ensure you set this in your .env file
SMTP_SERVER = 'smtp.office365.com'
SMTP_PORT = 587

# Function to validate email format
def is_valid_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email) is not None

# Function to create a visually appealing PDF ticket
def create_ticket_pdf(name, token, date, folder_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # Set background color for a gaming look
    pdf.set_fill_color(0, 0, 0)  # Black background
    pdf.rect(0, 0, 210, 297, 'F')  # A4 size

    # Title with a gaming font and color
    pdf.set_font("Arial", 'B', 24)
    pdf.set_text_color(255, 255, 255)  # White text
    pdf.cell(0, 20, "GamingCraft Hackathon 2024", ln=True, align='C')

    # Divider line
    pdf.set_draw_color(255, 0, 255)  # Magenta line for styling
    pdf.line(10, 30, 200, 30)

    # Participant name
    pdf.set_font("Arial", 'B', 16)
    pdf.set_text_color(0, 255, 255)  # Cyan text
    pdf.cell(0, 20, f"Name: {name}", ln=True, align='C')

    # Token and date
    pdf.set_font("Arial", 'B', 18)
    pdf.set_text_color(255, 102, 255)  # Light pink text
    pdf.cell(0, 20, f"Token Number: {token}", ln=True, align='C')
    pdf.set_text_color(255, 204, 0)  # Gold text for date
    pdf.cell(0, 20, f"Date: {date}", ln=True, align='C')

    # Add additional event information
    pdf.set_font("Arial", 'I', 12)
    pdf.set_text_color(200, 200, 200)
    pdf.cell(0, 20, "Join us for an exciting hackathon and be crowned the Gaming Champion!", ln=True, align='C')

    # Save the PDF file in the specified folder
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    pdf_file_name = os.path.join(folder_path, f"{name.replace(' ', '_')}_ticket.pdf")
    pdf.output(pdf_file_name)
    return pdf_file_name

# Load your Excel file
df = pd.read_excel('participants.xlsx')  # Adjust the path as necessary

# Initialize the starting token number
starting_token = 1000

# Email sending function
def send_email(name, email, token, pdf_file_name):
    message = MIMEMultipart()
    message['From'] = EMAIL_ADDRESS
    message['To'] = email
    message['Subject'] = 'Your Hackathon Participation Token'

    group_link = "https://linktr.ee/klef.mayavi"  # Replace with your actual Telegram group link
    body = f"""Hi {name},

Thank you for registering for our Hackathon!

Your unique participation token is: {token}
Please find your ticket details attached.

If you have any questions, join our group for updates: {group_link}

We look forward to your participation!

Best regards,
Team Mayavi
"""
    message.attach(MIMEText(body, 'plain'))

    # Attach the PDF file
    with open(pdf_file_name, 'rb') as pdf_file:
        part = MIMEApplication(pdf_file.read(), Name=os.path.basename(pdf_file_name))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(pdf_file_name)}"'
        message.attach(part)

    # Set up the server and send the email
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # Enable security
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, email, message.as_string())
            print(f"Email sent to {email}")
            return True  # Indicate success
    except Exception as e:
        print(f"Failed to send email to {email}: {e}")
        return False  # Indicate failure

# Generate tokens, create PDFs, and send emails
hackathon_date = "8-11-2024 to 9-11-2024"  # Set hackathon date
pdf_folder_path = 'tickets'  # Define the folder where PDFs will be saved

failed_emails = []  # List to hold participants who failed to receive emails

for index, row in df.iterrows():
    token = starting_token + index  # Increment token for each participant
    df.at[index, 'Token'] = token  # Assign unique token to the participant

    if is_valid_email(row['email']):  # Check email format
        pdf_file_name = create_ticket_pdf(row['name'], token, hackathon_date, pdf_folder_path)  # Create PDF ticket
        if not send_email(row['name'], row['email'], token, pdf_file_name):  # Send email with the token and PDF
            df.at[index, 'EmailSent'] = 'Failed'  # Mark email as failed
            failed_emails.append(row['name'])  # Add to failed email list
        else:
            df.at[index, 'EmailSent'] = 'Yes'  # Mark email as sent
    else:
        print(f"Invalid email format for {row['name']}: {row['email']}")

# Save changes back to the Excel file
df.to_excel('participants_with_tokens.xlsx', index=False)  # Adjust the path as necessary
print("Tokens generated and emails sent.")

# Create a new Excel file for failed emails
if failed_emails:
    failed_df = df[df['EmailSent'] == 'Failed']
    failed_df.to_excel('failed_email_participants.xlsx', index=False)  # Save failed emails to a new Excel file
    print("Failed emails saved to 'failed_email_participants.xlsx'.")

# Update JSON file with new token information
json_data = df.to_dict(orient='records')  # Convert DataFrame to a list of dictionaries
with open('participants.json', 'w') as json_file:
    json.dump(json_data, json_file, indent=4)  # Write JSON data to file
print("JSON file updated with participant data.")
