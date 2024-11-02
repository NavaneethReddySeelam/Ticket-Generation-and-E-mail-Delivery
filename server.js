const express = require('express');
const PDFDocument = require('pdfkit');
const nodemailer = require('nodemailer');
const fs = require('fs').promises;
const path = require('path');
const dotenv = require('dotenv');
const cors = require('cors');
const XLSX = require('xlsx');

dotenv.config();
const app = express();
const participants = require('./participants.json');

app.use(cors());
app.use(express.json());

const startToken = 1000;
const tokensUsed = new Set(participants.map(p => p.Token));
let currentToken = startToken;

while (tokensUsed.has(currentToken)) {
    currentToken++;
}

// Assign tokens if not present
participants.forEach(participant => {
    if (!participant.Token) {
        participant.Token = currentToken++;
    }
});

// Route to get participant details and send email
app.get('/getParticipant', async (req, res) => {
    const { token, email } = req.query;

    if (!token && !email) {
        return res.status(400).json({ error: 'Token or email must be provided' });
    }

    const participant = participants.find(p => (token && p.Token === parseInt(token)) || (email && p.email === email));

    if (!participant) {
        return res.status(404).json({ error: 'Participant not found' });
    }

    if (participant.emailSent) {
        return res.status(400).json({ error: 'Email has already been sent to this participant' });
    }

    try {
        const pdfBuffer = await generatePDF(participant);
        await sendEmail(participant.name, participant.email, participant.Token, pdfBuffer);
        participant.emailSent = true;

        // Update participants.json
        await fs.writeFile('./participants.json', JSON.stringify(participants, null, 2));
        res.status(200).json({ message: 'Email sent successfully', participant });
    } catch (error) {
        console.error(error);
        await logFailedEmail(participant);
        res.status(500).json({ error: 'Failed to send email, but participant found' });
    }
});

// Function to generate PDF
async function generatePDF(participant) {
    return new Promise((resolve, reject) => {
        const doc = new PDFDocument();
        const buffers = [];

        doc.on('data', buffers.push.bind(buffers));
        doc.on('end', () => resolve(Buffer.concat(buffers)));

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

// Function to send email with PDF
async function sendEmail(name, email, token, pdfBuffer) {
    const transporter = nodemailer.createTransport({
        host: 'smtp.office365.com',
        port: 587,
        secure: false,
        auth: {
            user: process.env.EMAIL_ADDRESS,
            pass: process.env.EMAIL_PASSWORD,
        },
    });

    const mailOptions = {
        from: process.env.EMAIL_ADDRESS,
        to: email,
        subject: 'Your Hackathon Participation Token',
        text: `Hello ${name},\n\nThank you for participating! Your token is ${token}. Please find your ticket attached.`,
        attachments: [
            {
                filename: `${name.replace(/ /g, '_')}_ticket.pdf`,
                content: pdfBuffer,
            },
        ],
    };

    return transporter.sendMail(mailOptions);
}

// Log failed emails uniquely
async function logFailedEmail(participant) {
    const failedEmailsPath = path.join(__dirname, 'failed_email_participants.json');
    const failedParticipants = await fs.readFile(failedEmailsPath)
        .then(data => JSON.parse(data))
        .catch(() => []);

    // Check if participant is already in failed list
    if (!failedParticipants.some(fp => fp.email === participant.email)) {
        failedParticipants.push(participant);
        await fs.writeFile(failedEmailsPath, JSON.stringify(failedParticipants, null, 2));
    }
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
