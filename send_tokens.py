# main.py
import os
import re
import json
import logging
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Set up logging
logging.basicConfig(level=logging.INFO)

# Set email credentials
EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.office365.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))

# Function to validate email format
def is_valid_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email) is not None

# Function to create custom ticket image
def create_custom_ticket_image(name, university_id, year, token, template_path, image_folder_path):
    try:
        template_image = Image.open(template_path).convert("RGBA")
    except FileNotFoundError:
        logging.error("Template image not found.")
        return None
    
    draw = ImageDraw.Draw(template_image)
    font_path = "D:\\hackathon\\CASTELAR.TTF"
    font_size_title = 70
    font_size_detail = 60

    try:
        font_title = ImageFont.truetype(font_path, font_size_title)
        font_detail = ImageFont.truetype(font_path, font_size_detail)
    except OSError:
        logging.error("Font file not found.")
        font_title = font_detail = ImageFont.load_default()

    token_text = f"Token: {token}"
    name_text = f"NAME: {name}"
    university_text = f"ID NO.: {university_id}"
    year_text = f"YEAR: {year}"

    texts = [token_text, name_text, university_text, year_text]
    image_width, image_height = template_image.size
    line_spacing = 40
    start_x = int(image_width * 0.20)
    start_y = int(image_height * 0.50)
    y = start_y

    for text in texts:
        font = font_title if text == name_text else font_detail
        text_bbox = draw.textbbox((0, 0), text, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        x = start_x + (image_width - start_x - text_width) // 2
        draw.text((x, y), text, font=font, fill="white")
        y += text_bbox[3] - text_bbox[1] + line_spacing

    output_image_path = os.path.join(image_folder_path, f"{name.replace(' ', '_')}_ticket.png")
    template_image.save(output_image_path)
    return output_image_path

# Function to send email with attached ticket image
def send_email(name, email, token, image_file_name):
    message = MIMEMultipart()
    message['From'] = EMAIL_ADDRESS
    message['To'] = email
    message['Subject'] = 'Your Hackathon Participation Token'
    body = f"Hi {name},\n\nYour token is: {token}.\n\nBest regards,\nTeam"
    message.attach(MIMEText(body, 'plain'))

    with open(image_file_name, 'rb') as image_file:
        part = MIMEApplication(image_file.read(), Name=os.path.basename(image_file_name))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(image_file_name)}"'
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
    participants_df = pd.read_excel('participants.xlsx')
    image_folder_path = 'D:\\hackathon\\tickets'
    template_path = "D:\\hackathon\\ticket.png"

    if not os.path.exists(image_folder_path):
        os.makedirs(image_folder_path)

    for index, row in participants_df.iterrows():
        name, email = row['name'], row['email']
        if is_valid_email(email):
            token = 1000 + index
            image_file_name = create_custom_ticket_image(name, row['universityId'], row['Year'], token, template_path, image_folder_path)
            if image_file_name and send_email(name, email, token, image_file_name):
                logging.info(f"Processed {email} successfully.")
            else:
                logging.error(f"Failed to process {email}.")
        else:
            logging.error(f"Invalid email for {name}: {email}")

if __name__ == "__main__":
    main()
