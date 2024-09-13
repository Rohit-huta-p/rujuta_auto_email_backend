const express = require('express')
const app = express();
const multer = require('multer');
const cron = require('node-cron')
const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');
const cors = require('cors')
app.use(cors())
const sendMail = require('./sendMail');
const upload = multer({ dest: 'uploads/' }); // Save to 'uploads/' directory



const clients = path.join(__dirname, 'uploads', 'contacts.xlsx'); 
const attachmentPath = path.join(__dirname, 'assets', 'contacts_2.xlsx');
const asigneedetails = path.join(__dirname, 'assets', 'asigneedetails.xlsx');


function processFile(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

  // Access data from the previous file
  if (previousFileData) {
    console.log("Previous file data:", previousFileData);
  }

  // Update previousFileData with the current file data
  previousFileData = sheetData;

  console.log("Current file data:", sheetData);
}
const emails = {
  'Rujuta': "r.a.sachin.kulkarni@accenture.com",
  'Rohit': "vcrohithuta@gmail.com"
};


const sendEmailsFromContacts = async () => {
  try {
    // Load the Excel file for contacts
    const contactsWorkbook = xlsx.readFile(clients);
    const contactsSheetName = contactsWorkbook.SheetNames[0]; 
    const contactsWorksheet = contactsWorkbook.Sheets[contactsSheetName];
    const contacts = xlsx.utils.sheet_to_json(contactsWorksheet);
 
    
    // Load the Excel file for ticket status (contacts_2.xlsx)
    const ticketStatusWorkbook = xlsx.readFile(attachmentPath); 
    const ticketSheetName = ticketStatusWorkbook.SheetNames[0]; 
    const ticketWorksheet = ticketStatusWorkbook.Sheets[ticketSheetName]; 
    const content = xlsx.utils.sheet_to_json(ticketWorksheet);

    
 
    const headers = Object.keys(content[0]);
    

    // Track if any "Last updated date" is empty
    let isAllUpdated = true;
    let assigneeEmails = {};
    const emails = {
      'Rujuta': "vchutagonna@gmail.com",
      'Rohit': "vcrohithuta.p@gmail.com"
    };
    console.log(emails);
  console.log(content);
  
  for (const row of content) {
    const lastUpdatedDate = row['Last updated date'];
    const assignee = row['Assignee'];

    if (!lastUpdatedDate) {
      isAllUpdated = false;

      // Group emails by assignee
      if (!assigneeEmails[assignee]) {
        assigneeEmails[assignee] = [];
      }

      // Assign email based on assignee
      if (assignee === 'Rujuta') {
        assigneeEmails[assignee] = emails.Rujuta;
      } else if (assignee === 'Rohit') {
        assigneeEmails[assignee] = emails.Rohit;
      }
    }
  }
console.log("fasdfsdf", assigneeEmails);
    
    if (isAllUpdated) {
      // If all rows are updated, send the regular email
      for (const contact of contacts) {
        const email = contact.emails;
        const name = contact.name;
        const subject = `Hello ${name}, This is a Test Email`;
        const message = `
          <h5>Dear ${name},\n\nThis is Ticket Status of today </h5>
          <table border='1'>
            <thead>
              ${headers.map(heading => `<th>${heading}</th>`).join('')}
            </thead>
            <tbody>
              ${content
                .map(row => `
                  <tr>
                    ${headers.map(header => `<td>${row[header] || ''}</td>`).join('')}
                  </tr>
                `).join('')}
            </tbody>
          </table>
          <br />
          <p>Best regards, Your Company</p>
        `;
        await sendMail(email, subject, message, attachmentPath);
      }
    } else {
      // Send a notification email to assignees with incomplete rows
      for (const [assignee, emaill] of Object.entries(assigneeEmails)) {


        
          const subject = `Pending Updates`;
          const message = `
            <h5>Dear ${assignee},</h5>
            <p>You have pending updates for the following rows:</p>
           
            <p>Please update these records as soon as possible.</p>
            <br />
            <p>Best regards, Your Company</p>
          `;
          await sendMail(emaill, subject, message, attachmentPath);
        
      }
    }
  } catch (error) {
    console.error('Error processing file or sending emails:', error);
  }
};

// Helper function to convert Excel date serial to JS date
function excelDateToJSDate(serial) {
  const excelStartDate = new Date(1899, 11, 30);
  return new Date(excelStartDate.getTime() + serial * 86400000).toLocaleDateString();
}

sendEmailsFromContacts()
app.listen(8001, () => {
    console.log("Server Started...");
    
})


