const express = require('express');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const xlsx = require('xlsx');
const cors = require('cors');
const http = require('http');
const cron = require('node-cron');
require('dotenv').config();

const app = express();
app.use(cors({ origin: '*' }));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

const httpServer = http.createServer(app);

const createExcelFile = (folderDetails) => {
    const workbook = xlsx.utils.book_new();

    folderDetails.forEach((folder) => {
        const worksheetData = folder.files.map(file => ({
            'File Name': file.name,
            'File Size (Bytes)': file.size,
            'Last Modified': file.lastModified
        }));

        worksheetData.unshift({
            'Folder Name': folder.name,
            'Total Size (Bytes)': folder.size
        });

        const worksheet = xlsx.utils.json_to_sheet(worksheetData);
        xlsx.utils.book_append_sheet(workbook, worksheet, folder.name);
    });

    const filePath = path.join(__dirname, `folder_details_${Date.now()}.xlsx`);
    xlsx.writeFile(workbook, filePath);
    return filePath;
};


const getFolderSizes = async (dir) => {
    let folderDetails = [];

    const items = await fs.promises.readdir(dir, { withFileTypes: true });

    for (const item of items) {
        const itemPath = path.join(dir, item.name);
        try {
            if (item.isDirectory()) {
                const size = await calculateFolderSize(itemPath);
                const files = await getFilesDetails(itemPath);
                folderDetails.push({ name: item.name, size, files });
            }
        } catch (err) {
            console.error(`Error accessing ${itemPath}:`, err);
        }
    }

    return folderDetails;
};


const calculateFolderSize = async (folderPath) => {
    let totalSize = 0;

    const items = await fs.promises.readdir(folderPath, { withFileTypes: true });

    for (const item of items) {
        const itemPath = path.join(folderPath, item.name);
        try {
            const stats = await fs.promises.stat(itemPath);

            if (stats.isDirectory()) {
                // Recursively calculate size of subdirectories
                totalSize += await calculateFolderSize(itemPath);
            } else {
                totalSize += stats.size;
            }
        } catch (err) {
            console.error(`Error accessing ${itemPath}:`, err);
        }
    }

    return totalSize;
};


const getFilesDetails = async (folderPath) => {
    const files = fs.readdirSync(folderPath, { withFileTypes: true });
    let filesDetails = [];

    for (const file of files) {
        const filePath = path.join(folderPath, file.name);
        try {
            const stats = await fs.promises.stat(filePath); // Use fs.promises.stat for asynchronous operation
            filesDetails.push({
                name: file.name,
                size: stats.size,
                lastModified: stats.mtime.toLocaleString() // Date and time of last modification
            });
        } catch (err) {
            console.error(`Error accessing ${filePath}:`, err);
        }
    }

    return filesDetails;
};

// Helper function to send email
const sendEmail = async (filePath) => {
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASSWORD
        },
        tls: {
            rejectUnauthorized: false
        }
    });

    const mailOptions = {
        from: process.env.EMAIL_USER,
        to: process.env.EMAIL_TO,
        subject: 'Folder Details Report',
        text: 'Attached is the folder details report.',
        attachments: [
            {
                filename: path.basename(filePath),
                path: filePath,
                contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }
        ]
    };

    try {
        await transporter.sendMail(mailOptions);
        console.log('Email sent successfully');
    } catch (error) {
        console.error('Error sending email:', error);
    }
};

// Function to execute the process
const executeProcess = async () => {
    const dir = process.env.USER_DIR;
    if (!dir) {
        console.error('USER_DIR is not set in .env file');
        return;
    }

    try {
        const folderDetails = await getFolderSizes(dir);
        const filePath = createExcelFile(folderDetails);
        await sendEmail(filePath);
        console.log('Email sent successfully');
    } catch (error) {
        console.error('An error occurred', error);
    }
};


// Schedule the task to run every minute
cron.schedule('* * * * *', () => {
    console.log('Running scheduled task');
    executeProcess();
});

const PORT = process.env.PORT || 3000;
httpServer.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});

module.exports = httpServer;
