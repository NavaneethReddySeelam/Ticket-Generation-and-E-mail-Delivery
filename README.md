
# Hackathon Ticket Generator and Email Sender

This project provides a script to generate custom tickets for hackathon participants and send them via email. 
The script ensures that each participant receives a unique ticket only once, and it retries sending for any 
failed attempts.

---

## Features

1. Generates a personalized ticket for each participant using provided details.
2. Sends emails with the ticket as an attachment.
3. Tracks participants in JSON format to avoid duplicate emails.
4. Logs failed and successful email attempts in Excel files.

---

## Requirements

1. **Install Dependencies**

   Make sure you have `pandas`, `Pillow`, and `python-dotenv` installed:
   ```bash
   pip install pandas pillow python-dotenv
   ```

2. **Environment Variables**

   Create a `.env` file in the project directory with your email credentials and SMTP server settings:
   ```env
   EMAIL_ADDRESS=your_email@example.com
   EMAIL_PASSWORD=your_password
   SMTP_SERVER=smtp.office365.com
   SMTP_PORT=587
   ```

3. **Required Files**

   - **Ticket Template**: Save the ticket template image as `ticket.png` and place it at the specified path (`D:\hackathon\ticket.png`).
   - **Font File**: Save the font file (e.g., `CASTELAR.TTF`) at the specified path (`D:\hackathon\CASTELAR.TTF`).
   - **Participants Data**: Prepare a `participants.xlsx` file with the following columns:
     - `name`: Participant's full name
     - `email`: Email address for sending the ticket
     - `universityId`: Unique university ID for the participant
     - `Year`: Academic year of the participant

---

## Usage

1. **Run the Script**

   Run the `main.py` script using the following command:
   ```bash
   python main.py
   ```

2. **Process Details**

   The script performs the following actions:
   - Reads participants' data from `participants.xlsx`.
   - Generates a ticket image for each participant.
   - Sends an email with the ticket attached.
   - Updates `failed_participants.xlsx` and `passed_participants.xlsx` based on success or failure.

3. **Log Files**
   - **failed_participants.xlsx**: Records participants for whom email sending failed.
   - **passed_participants.xlsx**: Records participants who successfully received the ticket.
   - **participants.json**: Maintains a list of participants who have already received the ticket.

---

## Troubleshooting

- **Missing Template or Font Files**: Ensure that `ticket.png` and `CASTELAR.TTF` are available at the specified paths.
- **SMTP Issues**: Check your email credentials and server details in the `.env` file.
- **Excel File Format**: Ensure `participants.xlsx` has the correct columns and valid data.

---

## License
CopyRights(c) Seelam Navaneeth Reddy
This project is licensed under the MIT License.
