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
const ticketsDir = path.join(__dirname, 'tickets');
const failedFilePath = './failed_participants.xlsx';
const passedFilePath = './successful_participants.xlsx';

let participants = [];
let failedParticipants = [];
let passedParticipants = [];

// Load participants from Excel file
async function loadParticipantsFromExcel() {
    try {
        const workbook = xlsx.readFile(participantsFilePath);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        participants = xlsx.utils.sheet_to_json(sheet);
        const tokensUsed = new Set(participants.map(p => p.Token).filter(Boolean));
        let currentToken = 1000;

        participants.forEach(participant => {
            if (!participant.Token) {
                while (tokensUsed.has(currentToken)) currentToken++;
                participant.Token = currentToken++;
                tokensUsed.add(participant.Token);
            }
        });
    } catch (error) {
        console.error('Failed to load participants from Excel:', error);
    }
}

// Load passed and failed participants from Excel files
async function loadStatusLists() {
    try {
        if (await fs.access(passedFilePath).then(() => true).catch(() => false)) {
            const workbook = xlsx.readFile(passedFilePath);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            passedParticipants = xlsx.utils.sheet_to_json(sheet);
        }

        if (await fs.access(failedFilePath).then(() => true).catch(() => false)) {
            const workbook = xlsx.readFile(failedFilePath);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            failedParticipants = xlsx.utils.sheet_to_json(sheet);
        }
    } catch (error) {
        console.error('Failed to load passed and failed lists:', error);
    }
}

// Append participant data to an Excel file
async function addToExcelFile(filePath, participant) {
    try {
        let workbook;
        if (await fs.access(filePath).then(() => true).catch(() => false)) {
            workbook = xlsx.readFile(filePath);
        } else {
            workbook = xlsx.utils.book_new();
        }

        const sheetName = 'Participants';
        const dataSheet = workbook.Sheets[sheetName] || xlsx.utils.json_to_sheet([]);
        const data = xlsx.utils.sheet_to_json(dataSheet);

        data.push(participant);
        workbook.Sheets[sheetName] = xlsx.utils.json_to_sheet(data);
        xlsx.writeFile(workbook, filePath);
    } catch (error) {
        console.error('Failed to update Excel file:', error);
    }
}

// Check if participant is in a list by university ID
function isInList(list, universityId) {
    return list.some(p => p.UniversityId === universityId);
}

// Send email with PNG attachment
async function sendEmail(participant, pngPath) {
    if (!process.env.SMTP_HOST || !process.env.EMAIL_ADDRESS || !process.env.EMAIL_PASSWORD) {
        console.error('Missing required environment variables for email configuration');
        return false;
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
        return true;
    } catch (error) {
        console.error(`Failed to send email to ${participant.Email}: ${error}`);
        return false;
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

// Process participants and update lists based on email success or failure
async function processParticipants() {
    for (const participant of participants) {
        if (isInList(passedParticipants, participant.UniversityId)) {
            // Skip if already in passed list
            console.log(`Skipping ${participant.Name}, email already sent.`);
            continue;
        }

        const pngPath = await createTicketImage(participant);
        const emailSuccess = await sendEmail(participant, pngPath);

        if (emailSuccess) {
            await addToExcelFile(passedFilePath, participant);
            passedParticipants.push(participant);
            failedParticipants = failedParticipants.filter(p => p.UniversityId !== participant.UniversityId);
        } else {
            await addToExcelFile(failedFilePath, participant);
            failedParticipants.push(participant);
        }

        await fs.unlink(pngPath); // Delete PNG after sending
    }
}

// Endpoint to start the email process
app.get('/sendTickets', async (req, res) => {
    try {
        await loadParticipantsFromExcel();
        await loadStatusLists();
        await processParticipants();
        res.status(200).send("Emails processed successfully!");
    } catch (error) {
        console.error('Error processing participants:', error);
        res.status(500).send("Failed to process participants.");
    }
});

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
