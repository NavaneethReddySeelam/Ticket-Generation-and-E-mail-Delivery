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

# Paths and constants
TEMPLATE_PATH = "D:\\hackathon\\ticket.png"
FONT_PATH = "D:\\hackathon\\CASTELAR.TTF"
IMAGE_FOLDER_PATH = "D:\\hackathon\\tickets"
PARTICIPANTS_JSON = "participants.json"
FAILED_PARTICIPANTS_FILE = "failed_participants.xlsx"
PASSED_PARTICIPANTS_FILE = "passed_participants.xlsx"

# Load participants from JSON or initialize empty list
def load_participants():
    if os.path.exists(PARTICIPANTS_JSON):
        with open(PARTICIPANTS_JSON, 'r') as f:
            return json.load(f)
    return []

# Save participants to JSON
def save_participants(participants):
    with open(PARTICIPANTS_JSON, 'w') as f:
        json.dump(participants, f, indent=4)

# Validate email format
def is_valid_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email) is not None

# Create ticket image
def create_custom_ticket_image(name, university_id, year, token):
    try:
        template_image = Image.open(TEMPLATE_PATH).convert("RGBA")
    except FileNotFoundError:
        logging.error("Template image not found.")
        return None
    
    draw = ImageDraw.Draw(template_image)
    try:
        font_title = ImageFont.truetype(FONT_PATH, 70)
        font_detail = ImageFont.truetype(FONT_PATH, 60)
    except OSError:
        logging.error("Font file not found.")
        font_title = font_detail = ImageFont.load_default()

    texts = [
        f"Token: {token}",
        f"NAME: {name}",
        f"ID NO.: {university_id}",
        f"YEAR: {year}"
    ]
    image_width, image_height = template_image.size
    start_x = int(image_width * 0.20)
    start_y = int(image_height * 0.50)
    y = start_y

    for text in texts:
        font = font_title if "NAME" in text else font_detail
        text_bbox = draw.textbbox((0, 0), text, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        x = start_x + (image_width - start_x - text_width) // 2
        draw.text((x, y), text, font=font, fill="white")
        y += text_bbox[3] - text_bbox[1] + 40

    output_image_path = os.path.join(IMAGE_FOLDER_PATH, f"{name.replace(' ', '_')}_ticket.png")
    os.makedirs(IMAGE_FOLDER_PATH, exist_ok=True)
    template_image.save(output_image_path)
    return output_image_path

# Send email with attachment
def send_email(name, email, token, image_file_name, smtp_server):
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
        smtp_server.sendmail(EMAIL_ADDRESS, email, message.as_string())
        logging.info(f"Email sent to {email}")
        return True
    except Exception as e:
        logging.error(f"Failed to send email to {email}: {e}")
        return False

# Load or create status list
def load_status_list(file_path):
    if os.path.exists(file_path):
        return pd.read_excel(file_path)
    else:
        return pd.DataFrame(columns=["name", "email", "universityId", "Year", "Token"])

# Save status list to file
def save_status_list(df, file_path):
    df.to_excel(file_path, index=False)

# Main process
def main():
    participants_df = pd.read_excel('participants.xlsx')
    failed_list = load_status_list(FAILED_PARTICIPANTS_FILE)
    passed_list = load_status_list(PASSED_PARTICIPANTS_FILE)
    participants = load_participants()

    # Prepare SMTP server connection
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp_server:
        smtp_server.starttls()
        smtp_server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

        for index, row in participants_df.iterrows():
            name, email = row['name'], row['email']
            university_id, year = row['universityId'], row['Year']
            
            # Skip if already processed
            if any(passed_list["email"] == email) or any(failed_list["email"] == email):
                logging.info(f"Skipping {name} ({email}) - Already processed.")
                continue
            
            # Validate email and create ticket
            if is_valid_email(email):
                token = 1000 + index
                image_file_name = create_custom_ticket_image(name, university_id, year, token)
                
                # Send email
                if image_file_name and send_email(name, email, token, image_file_name, smtp_server):
                    row['Token'] = token
                    passed_list = passed_list.append(row, ignore_index=True)
                    logging.info(f"Processed {email} successfully.")
                else:
                    failed_list = failed_list.append(row, ignore_index=True)
                    logging.error(f"Failed to process {email}.")
            else:
                failed_list = failed_list.append(row, ignore_index=True)
                logging.error(f"Invalid email for {name}: {email}")

    # Save updated lists
    save_status_list(failed_list, FAILED_PARTICIPANTS_FILE)
    save_status_list(passed_list, PASSED_PARTICIPANTS_FILE)
    
    # Update JSON
    save_participants(participants)

if __name__ == "__main__":
    main()
