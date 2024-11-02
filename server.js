const express = require('express');
const PDFDocument = require('pdfkit');
const nodemailer = require('nodemailer');
const fs = require('fs').promises; // Use promises for file operations
const path = require('path');
const dotenv = require('dotenv');
const cors = require('cors');
const XLSX = require('xlsx');

dotenv.config(); // Load environment variables from .env file

const app = express();
const participants = require('./participants.json');

app.use(cors()); // Enable CORS
app.use(express.json());

// Generate unique token numbers starting from 1000
const startToken = 1000;
const tokensUsed = new Set(participants.map(p => p.Token)); // Track used tokens
let currentToken = startToken;

while (tokensUsed.has(currentToken)) {
    currentToken++;
}

// Assign unique token to each participant if not already assigned
participants.forEach((participant) => {
    if (!participant.Token) {
        participant.Token = currentToken++;
    }
});

// Endpoint to get participant details and optionally send an email
app.get('/getParticipant', async (req, res) => {
    const token = parseInt(req.query.token, 10);
    const email = req.query.email;

    if (!token && !email) {
        return res.status(400).json({ error: 'Token or email must be provided' });
    }

    const participantIndex = participants.findIndex(
        p => (token && p.Token === token) || (email && p.email === email)
    );

    if (participantIndex === -1) {
        return res.status(404).json({ error: 'Participant not found' });
    }

    const participant = participants[participantIndex];

    if (participant.emailSent) {
        return res.status(400).json({ error: 'Email has already been sent to this participant' });
    }

    try {
        const pdfBuffer = await generatePDF(participant);
        await sendEmail(participant.name, participant.email, participant.Token, pdfBuffer);
        participants[participantIndex].emailSent = true;

        // Save updated participants array back to JSON asynchronously
        await fs.writeFile('./participants.json', JSON.stringify(participants, null, 2));

        res.status(200).json({ message: 'Email sent successfully', participant });
    } catch (error) {
        console.error(error);
        await logFailedEmail(participant); // Log to Excel if email fails
        res.status(500).json({ error: 'Email not sent, but participant found' });
    }
});

// Function to generate PDF and return as buffer
function generatePDF(participant) {
    return new Promise((resolve, reject) => {
        const doc = new PDFDocument();
        const buffers = [];

        doc.on('data', buffers.push.bind(buffers));
        doc.on('end', () => {
            const pdfBuffer = Buffer.concat(buffers);
            resolve(pdfBuffer);
        });

        // Add content to PDF
        doc.fontSize(25).text('GameCraft Hackathon 2024', { align: 'center' });
        doc.moveDown();
        doc.fontSize(16).text(`Name: ${participant.name}`, { underline: true });
        doc.text(`University ID: ${participant.universityId || 'N/A'}`);
        doc.text(`Year: ${participant.year || 'N/A'}`);
        doc.moveDown();
        doc.text(`Token Number: ${participant.Token}`, { bold: true });
        doc.text(`Date: 8-11-2024 to 9-11-2024`);
        doc.end();
    });
}

// Function to send email with PDF attachment
function sendEmail(name, email, token, pdfBuffer) {
    const transporter = nodemailer.createTransport({
        host: 'smtp.office365.com',
        port: 587,
        secure: false,
        auth: {
            user: process.env.EMAIL_ADDRESS,  // Get email from .env
            pass: process.env.EMAIL_PASSWORD,  // Get password from .env
        },
    });

    const mailOptions = {
        from: process.env.EMAIL_ADDRESS, // Use environment variable
        to: email,
        subject: 'Your Hackathon Participation Token',
        text: `Hi ${name},\n\nThank you for registering for GameCraft Hackathon 2024!\n\nYour unique participation token is: ${token}\n\nPlease find your ticket details attached.\n\nBest regards,\nThe Hackathon Team`,
        attachments: [
            {
                filename: 'ticket.pdf',
                content: pdfBuffer,
                contentType: 'application/pdf',
            },
        ],
    };

    return new Promise((resolve, reject) => {
        transporter.sendMail(mailOptions, (error, info) => {
            if (error) {
                console.error('Error sending email:', error);
                return reject(error);
            }
            console.log('Email sent:', info.response);
            resolve(true);
        });
    });
}

// Function to log failed emails to an Excel file
async function logFailedEmail(participant) {
    const filePath = path.join(__dirname, 'failed_participants.xlsx');

    // Create a new workbook or load the existing one
    const failedParticipants = fs.existsSync(filePath) ? XLSX.readFile(filePath) : XLSX.utils.book_new();
    const ws = failedParticipants.Sheets['Failed Emails'] || XLSX.utils.aoa_to_sheet([[
        'Name',
        'University ID',
        'Email',
        'Category',
        'Year',
    ]]);

    // Add participant's details to the failed list
    const newRow = [
        participant.name,
        participant.universityId,
        participant.email,
        participant.category || 'N/A',
        participant.year || 'N/A',
    ];

    // Append new row to the sheet
    const currentData = XLSX.utils.sheet_to_json(ws, { header: 1 });
    currentData.push(newRow);
    const newWs = XLSX.utils.aoa_to_sheet(currentData);
    failedParticipants.Sheets['Failed Emails'] = newWs;

    // Write updated Excel file
    XLSX.writeFile(failedParticipants, filePath);
    console.log(`Logged failed email for ${participant.name} to ${filePath}`);
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
