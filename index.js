const express = require('express')
const app = express();
const multer = require('multer');
const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');
const cors = require('cors')
app.use(cors())
const sendMail = require('./sendMail');
const upload = multer({ dest: 'uploads/' }); // Save to 'uploads/' directory



const clients = path.join(__dirname, 'uploads', 'contacts.xlsx'); 
const attachmentPath = path.join(__dirname, 'assets', 'contacts_2.xlsx');

// Email sending function
const sendEmailsFromContacts = async () => {
  try {
  // Load the Excel file for contacts
  const contactsWorkbook = xlsx.readFile(clients);
  const contactsSheetName = contactsWorkbook.SheetNames[0]; // Get the first sheet
  const contactsWorksheet = contactsWorkbook.Sheets[contactsSheetName];
  const contacts = xlsx.utils.sheet_to_json(contactsWorksheet);
  console.log("contacts", contacts);

  // Load the Excel file for ticket status (contacts_2.xlsx)
  const ticketStatusWorkbook = xlsx.readFile(attachmentPath); // Use 'ticketStatusWorkbook' instead of 'workbook'
  const ticketSheetName = ticketStatusWorkbook.SheetNames[0]; // Get the first sheet
  const ticketWorksheet = ticketStatusWorkbook.Sheets[ticketSheetName]; // Correct reference to 'ticketStatusWorkbook'
  const content = xlsx.utils.sheet_to_json(ticketWorksheet);
  const headers = Object.keys(content[0]);

    console.log("contacfsdfsdft_2",content);

    // console.log("Headers",headers);
    
    


    // Iterate through contacts and send emails
    for (const contact of contacts) {
      const email = contact.emails; // Replace with your actual column name in the Excel sheet
      const name = contact.name; // Replace with your actual column name
      console.log(headers);
      
      // Customize the email content
      const subject = `Hello ${name}, This is a Test Email`;
      const message = `
        <h5>Dear ${name},\n\nThis is Ticket Status of today </h5>
        <table border='1'>
          <thead>
         
            ${headers.map(heading => `<th>${heading}</th>`).join('')}
          
          </thead>
          <tbody>
          ${content
            .filter(row => Object.values(row).some(value => value)) // Filter out empty rows
            .map(row => (
              `<tr>
                ${headers.map(header => `<td>${row[header] || ''}</td>`).join('')} 
              </tr>`
            )).join('')}
          </tbody>
        </table>
        <br />
        <br />
        <p>Best regards, Your Company</p>
      `;

      // Send email to each contact with the attachment
      await sendMail(email, subject, message, attachmentPath);
    }
  } catch (error) {
    console.error('Error processing file or sending emails:', error);
  }
};

cron.schedule('0 9 * * *', () => {
  console.log('Sending daily emails at 4:21 PM');
  sendEmailsFromContacts();
});




app.listen(8001, () => {
    console.log("Server Started...");
    
})


