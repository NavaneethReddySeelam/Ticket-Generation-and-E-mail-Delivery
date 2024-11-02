readme_content = """# **Event Ticket Generator and Email Dispatcher**

## Overview

The **Event Ticket Generator and Email Dispatcher** is a web application designed to manage participant registrations for various events. This project generates unique participation tokens, creates visually appealing PDF tickets, and sends confirmation emails to participants. The system tracks participants' registration status and stores relevant data for further analysis.

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
   git clone git@github.com:NavaneethReddySeelam/Ticket-Generation-and-E-mail-Delivery.git
   cd Ticket-Generation-and-E-mail-Delivery
   cd event_ticket_generator
