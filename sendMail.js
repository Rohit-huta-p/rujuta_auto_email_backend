const nodemailer = require('nodemailer');
const xlsx = require('xlsx');
const path = require('path');

// 3. Configure Nodemailer
let transporter = nodemailer.createTransport({
    service: 'gmail', // Replace with your email service
    auth: {
      user: 'vcrohithuta.p@gmail.com', // Replace with your email
      pass: 'xgjdgbkxlmwfwdbm'   // Replace with your email password or app password
    }
  });

// 4. Function to send emails
const sendEmail = async (recipient, subject, message, attachmentPath) => {
    let mailOptions = {
      from: 'vcrohithuta@gmail.com',
      to: recipient,
      subject: subject,
      text: message,
      attachments: [
        {
            filename: path.basename(attachmentPath),
            path: attachmentPath
        }
    ]
    };
  
    try {
      let info = await transporter.sendMail(mailOptions);
      console.log(`Email sent to ${recipient}: ${info.response}`);
    } catch (error) {
      console.error(`Failed to send email to ${recipient}:`, error);
    }
  };

module.exports = sendEmail;