
# Hackathon Ticket Generation and Email Delivery System

This project is a **Hackathon Registration and Ticketing System** designed to automate the process of generating unique participant tokens, creating personalized tickets, and sending these tickets as email attachments to participants. The system uses **Node.js, Express, Nodemailer, Canvas**, and **Excel** to manage participant data and send out custom ticket images.

---

## Table of Contents

- [Project Overview](#project-overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Project Structure](#project-structure)
- [Error Handling](#error-handling)
- [Future Improvements](#future-improvements)
- [Contributing](#contributing)
- [License](#license)

---

## Project Overview

This system automates the ticket generation process for a hackathon, including:
- Loading participant details from an Excel sheet.
- Generating a unique token number for each participant.
- Creating a personalized PNG ticket for each participant using Canvas.
- Sending an email with the ticket attached to each participant.

The token generation starts at 1000 and increments by one for each participant, ensuring unique identification. Invalid emails are skipped, and errors are logged.

## Features

- **Unique Token Generation**: Tokens start at 1000 and are incrementally assigned.
- **Personalized Ticket Creation**: Tickets are generated with participant-specific information.
- **Email Delivery**: Sends tickets as email attachments using Nodemailer.
- **Excel Support**: Participant data is loaded from and saved to Excel sheets.
- **Error Logging**: Logs errors for missing data or email send failures.

## Prerequisites

- **Node.js** and **npm**
- **Python 3.x**
- **pip** package manager
- **Excel file** containing participant data

## Installation

1. **Clone the Repository**:
    ```bash
    git clone <repository-url>
    cd <repository-folder>
    ```

2. **Install Dependencies**:
    For Node.js server:
    ```bash
    npm install
    ```

    For Python script:
    ```bash
    pip install -r requirements.txt
    ```

3. **Directory Setup**:
   Ensure you have the following file paths:
   - `D:/hackathon/ticket.png` (Template image for tickets)
   - `D:/hackathon/CASTELAR.TTF` (Font file for tickets)

## Configuration

1. **Environment Variables**: Create a `.env` file in the project root directory and add the following:
   ```plaintext
   SMTP_HOST=smtp.your-email-provider.com
   SMTP_PORT=587
   EMAIL_ADDRESS=your-email@example.com
   EMAIL_PASSWORD=your-email-password
   ```

   Ensure your `.env` file is **not tracked in Git** to keep credentials secure.

2. **Excel Sheet Format**: Ensure your `participants.xlsx` file contains columns:
   - `Name`, `UniversityId`, `Email`, `Category`, and `Year`.

## Usage

1. **Load and Generate Tokens**:
   - Run the server:
     ```bash
     node server.js
     ```
   - The server will load participants, generate tokens, and store updated data.

2. **Send Tickets**:
   - Endpoint to generate and send tickets:
     ```bash
     GET /sendTickets
     ```
   - The system creates a PNG ticket, attaches it to an email, and sends it to each participant.

3. **Python Script (Alternative)**:
   Run the `send_tokens.py` script to perform the same tasks as the server:
   ```bash
   python send_tokens.py
   ```

## Project Structure

```
.
├── server.js               # Main server file
├── send_tokens.py          # Python script for ticket generation
├── participants.xlsx       # Participant data (Excel format)
├── tickets/                # Folder for generated ticket images
├── .env                    # Environment variables
├── README.md               # Project documentation
└── package.json            # Node.js dependencies
```

## Error Handling

- **Email Send Failures**: Participants who do not receive emails due to errors are logged.
- **File Path Issues**: Errors for missing template image or font file are logged in the console.

## Future Improvements

- **Dynamic Template Selection**: Allow users to choose different ticket templates.
- **Enhanced Error Reporting**: Include more detailed error logs and email validation.
- **Database Integration**: Replace the Excel sheet with a database for improved scalability.

## Contributing

Feel free to contribute by opening a pull request.

## License

This project is licensed under the MIT License.
