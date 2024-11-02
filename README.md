# GameCraft Hackathon 2024 Registration System

## Overview
The **GameCraft Hackathon 2024** Registration System is a web application designed to manage participant registrations for the hackathon event. This project generates unique participation tokens, creates visually appealing PDF tickets, and sends confirmation emails to participants. The system tracks participants' registration status and stores relevant data for further analysis.

## Features
- **Participant Registration**: Collects participant details including name, university ID, email, year, and branch.
- **Token Generation**: Automatically generates unique participation tokens starting from 1000.
- **PDF Ticket Generation**: Creates personalized PDF tickets for each participant with their details and event information.
- **Email Notifications**: Sends confirmation emails with the ticket attached to participants using Outlook's SMTP service.
- **Data Management**: Stores participant information and updates status in an Excel sheet and JSON file.
- **Error Handling**: Logs failed email attempts and saves the details for follow-up.

## Technologies Used
- **Python**: Programming language used for backend development.
- **Pandas**: Library for data manipulation and analysis.
- **Smtplib**: Library for sending emails via SMTP.
- **FPDF**: Library for generating PDF documents.
- **Python-dotenv**: Package for loading environment variables from a `.env` file.
- **Openpyxl**: Library for reading and writing Excel files.
- **Regular Expressions**: Used for validating email formats.

## Setup Instructions
1. **Clone the Repository**:
   ```bash
   git clone https://github.com/NavaneethReddySeelam/GameCraft-Hackathon-2024.git
   cd GameCraft-Hackathon-2024
   ```

2. **Create a `.env` File**: 
   Create a file named `.env` in the root directory and add your email credentials:
   ```
   EMAIL_ADDRESS=your_email@example.com
   EMAIL_PASSWORD=your_password
   ```

3. **Install Required Packages**:
   You can use `pip` to install the required libraries. If you have a `requirements.txt` file, run:
   ```bash
   pip install -r requirements.txt
   ```

4. **Prepare Excel File**:
   Make sure you have an Excel file named `participants.xlsx` in the root directory with the following columns: `name`, `email`.

5. **Run the Script**:
   Execute the main script to start the registration process:
   ```bash
   python main.py
   ```

## Usage
- The script reads participant information from the `participants.xlsx` file.
- It generates unique tokens, creates PDF tickets, and sends emails.
- After execution, an updated Excel file `participants_with_tokens.xlsx` will be generated along with a JSON file `participants.json`.

## Error Handling
- Any failed email notifications will be logged in a new Excel file named `failed_email_participants.xlsx`.

## Contributing
Contributions are welcome! If you have suggestions for improvements or features, please open an issue or submit a pull request.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgments
- Special thanks to the libraries and resources that made this project possible.
