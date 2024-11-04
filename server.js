const express = require('express');
const nodemailer = require('nodemailer');
const fs = require('fs').promises;
const path = require('path');
const dotenv = require('dotenv');
const cors = require('cors');
const { createCanvas } = require('canvas');
const xlsx = require('xlsx');

dotenv.config();
const app = express();
app.use(cors());
app.use(express.json());

const participantsFilePath = './participants.xlsx';
const participantsJsonFilePath = './participants.json';
const ticketsDir = path.join(__dirname, 'tickets');
let participants = [];

// Load participants from JSON file
async function loadParticipants() {
    try {
        const data = await fs.readFile(participantsJsonFilePath, 'utf-8');
        participants = JSON.parse(data);
    } catch (error) {
        if (error.code !== 'ENOENT') {
            console.error('Failed to load participants:', error);
        }
        participants = [];
    }
}

// Save participants to JSON file with safe write operation
async function saveParticipants() {
    try {
        const tempFilePath = `${participantsJsonFilePath}.tmp`;
        await fs.writeFile(tempFilePath, JSON.stringify(participants, null, 2));
        await fs.rename(tempFilePath, participantsJsonFilePath);
    } catch (error) {
        console.error('Failed to save participants:', error);
    }
}

// Load participants from Excel file and generate unique tokens
async function loadParticipantsFromExcel() {
    try {
        const workbook = xlsx.readFile(participantsFilePath);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        const loadedParticipants = xlsx.utils.sheet_to_json(sheet);
        const tokensUsed = new Set(participants.map(p => p.Token));
        let currentToken = 1000;

        loadedParticipants.forEach(participant => {
            if (!participant.Token) {
                while (tokensUsed.has(currentToken)) currentToken++;
                participant.Token = currentToken++;
                tokensUsed.add(participant.Token);
            }
            if (!participants.some(p => p.UniversityId === participant.UniversityId)) {
                participants.push(participant);
            }
        });

        await saveParticipants();
    } catch (error) {
        console.error('Failed to load participants from Excel:', error);
    }
}

// Send email with PNG attachment
async function sendEmail(participant, pngPath) {
    if (!process.env.SMTP_HOST || !process.env.EMAIL_ADDRESS || !process.env.EMAIL_PASSWORD) {
        console.error('Missing required environment variables for email configuration');
        return;
    }

    const transporter = nodemailer.createTransport({
        host: process.env.SMTP_HOST,
        port: process.env.SMTP_PORT || 587,
        secure: process.env.SMTP_SECURE === 'true',
        auth: {
            user: process.env.EMAIL_ADDRESS,
            pass: process.env.EMAIL_PASSWORD,
        },
    });

    const mailOptions = {
        from: process.env.EMAIL_ADDRESS,
        to: participant.Email,
        subject: 'Your Hackathon Participation Token',
        text: `Hi ${participant.Name},\n\nThank you for registering for our Hackathon!\nYour unique participation token is: ${participant.Token}\nPlease find your ticket attached.\n\nBest regards,\nTeam Mayavi`,
        attachments: [
            {
                filename: `${participant.Name.replace(/\s+/g, '_')}_ticket.png`,
                path: pngPath,
            }
        ]
    };

    try {
        await transporter.sendMail(mailOptions);
        console.log(`Email sent to ${participant.Email}`);
    } catch (error) {
        console.error(`Failed to send email to ${participant.Email}: ${error}`);
    }
}

// Create ticket image with participant details
async function createTicketImage(participant) {
    const { Name, UniversityId, Year, Token } = participant;
    const width = 600;
    const height = 400;
    const canvas = createCanvas(width, height);
    const ctx = canvas.getContext('2d');

    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);
    ctx.fillStyle = '#000000';
    ctx.font = '28px Arial';

    const xOffset = width * 0.60;
    const startYOffset = height * 0.3;

    ctx.fillText(`Name: ${Name}`, xOffset, startYOffset);
    ctx.fillText(`University ID: ${UniversityId}`, xOffset, startYOffset + 50);
    ctx.fillText(`Year: ${Year}`, xOffset, startYOffset + 100);
    ctx.fillText(`Token: ${Token}`, xOffset, startYOffset + 150);

    await fs.mkdir(ticketsDir, { recursive: true });

    const outputPath = path.join(ticketsDir, `${Name.replace(/\s+/g, '_')}_ticket_${Date.now()}.png`);
    const buffer = canvas.toBuffer('image/png');
    await fs.writeFile(outputPath, buffer);

    return outputPath;
}

// Route to generate images and send emails to all participants
app.get('/sendTickets', async (req, res) => {
    try {
        await loadParticipantsFromExcel();

        for (const participant of participants) {
            if (!participant.Email || !isValidEmail(participant.Email)) {
                console.warn(`Invalid email for participant ${participant.Name}: ${participant.Email}`);
                continue;
            }
            const pngPath = await createTicketImage(participant);
            await sendEmail(participant, pngPath);
            participant.EmailSent = 'Yes';

            await fs.unlink(pngPath);
        }

        await saveParticipants();
        res.status(200).send("Emails sent successfully to all participants!");
    } catch (error) {
        console.error('Error processing participants:', error);
        res.status(500).send("Failed to process participants.");
    }
});

// Email validation function
function isValidEmail(email) {
    const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return regex.test(email);
}

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});

// Initialize by loading participants
loadParticipants();
