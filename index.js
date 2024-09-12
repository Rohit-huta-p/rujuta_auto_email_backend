const express = require('express')
const app = express();
const multer = require('multer');
const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');
const cors = require('cors')
app.use(cors())
const cron = require('node-cron');
const sendMail = require('./sendMail');
const upload = multer({ dest: 'uploads/' }); // Save to 'uploads/' directory

// app.post('/upload', upload.single('file'), async (req, res) => {
//     const file = req.file;
    
//     if (!file) {
//       return res.status(400).send('No file uploaded.');
//     }
  
//     const attachmentPath = path.join(__dirname, 'assets', 'contacts_2.xlsx'); // Path to the attachment

//     // Move file to a new location
//     const newLocation = path.join(__dirname, 'uploads', file.originalname);
//     fs.rename(file.path, newLocation, async (err) => {
//       if (err) {
//         return res.status(500).send('Error saving file.');
//       }
      
//       try {
//         // Load the Excel file
//         const workbook = xlsx.readFile(newLocation);
//         const sheetName = workbook.SheetNames[0]; // Get the first sheet
//         const worksheet = workbook.Sheets[sheetName];
//         const contacts = xlsx.utils.sheet_to_json(worksheet);
//         console.log(contacts);
        
//         // Iterate through contacts and send emails
//         for (const contact of contacts) {
//           const email = contact.emails; // Replace with your actual column name in the Excel sheet
//           const name = contact.name;   // Replace with your actual column name
        
//           // Customize the email content
//           const subject = `Hello ${name}, This is a Test Email`;
//           const message = `Dear ${name},\n\nThis is an automated message for email automation testing.\n\nBest regards,\nYour Company`;
        
//           // Send email to each contact
//           await sendMail(email, subject, message,attachmentPath );
//         }

//         res.send('File uploaded and emails sent.');
//       } catch (error) {
//         console.error('Error processing file or sending emails:', error);
//         res.status(500).send('Error processing file or sending emails.');
//       }
//     });
// });



const attachmentPath = path.join(__dirname, 'assets', 'contacts_2.xlsx'); // Path to the attachment

// Email sending function
const sendEmailsFromContacts = async () => {
  try {
    // Load the Excel file
    const workbook = xlsx.readFile(attachmentPath);
    const sheetName = workbook.SheetNames[0]; // Get the first sheet
    const worksheet = workbook.Sheets[sheetName];
    const contacts = xlsx.utils.sheet_to_json(worksheet);
    console.log(contacts);

    // Iterate through contacts and send emails
    for (const contact of contacts) {
      const email = contact.emails; // Replace with your actual column name in the Excel sheet
      const name = contact.name; // Replace with your actual column name

      // Customize the email content
      const subject = `Hello ${name}, This is a Test Email`;
      const message = `Dear ${name},\n\nThis is an automated message for email automation testing.\n\nBest regards,\nYour Company`;

      // Send email to each contact with the attachment
      await sendMail(email, subject, message, attachmentPath);
    }
  } catch (error) {
    console.error('Error processing file or sending emails:', error);
  }
};

// Schedule the email to be sent at 9:00 AM every day
cron.schedule('0 9 * * *', () => {
    console.log('Sending daily emails at 4:21 PM');
    sendEmailsFromContacts();
});





app.listen(8001, () => {
    console.log("Server Started...");
    
})


