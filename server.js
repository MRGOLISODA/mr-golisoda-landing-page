const express = require('express');
const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

const workbookPath = path.join(__dirname, 'leads.xlsx');
let workbook = new ExcelJS.Workbook();
let worksheet;

async function initWorkbook() {
  if (fs.existsSync(workbookPath)) {
    await workbook.xlsx.readFile(workbookPath);
    worksheet = workbook.getWorksheet('Leads');
    if (!worksheet) {
      worksheet = workbook.addWorksheet('Leads');
      worksheet.columns = [
        { header: 'Name', key: 'name', width: 20 },
        { header: 'Email', key: 'email', width: 30 },
        { header: 'Phone', key: 'phone', width: 15 },
        { header: 'City', key: 'city', width: 20 },
        { header: 'Message', key: 'message', width: 50 },
        { header: 'Date', key: 'date', width: 20 }
      ];
    }
  } else {
    worksheet = workbook.addWorksheet('Leads');
    worksheet.columns = [
      { header: 'Name', key: 'name', width: 20 },
      { header: 'Email', key: 'email', width: 30 },
      { header: 'Phone', key: 'phone', width: 15 },
      { header: 'City', key: 'city', width: 20 },
      { header: 'Message', key: 'message', width: 50 },
      { header: 'Date', key: 'date', width: 20 }
    ];
    await workbook.xlsx.writeFile(workbookPath);
  }
}

initWorkbook();

const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST || 'smtp.example.com',
  port: parseInt(process.env.SMTP_PORT || '587'),
  secure: false,
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS
  }
});

app.post('/api/lead', async (req, res) => {
  const { name, email, phone, city, message } = req.body;
  try {
    worksheet.addRow({
      name,
      email,
      phone,
      city,
      message,
      date: new Date().toISOString()
    });
    await workbook.xlsx.writeFile(workbookPath);

    const mailOptions = {
      from: process.env.SMTP_FROM || process.env.SMTP_USER,
      to: process.env.LEAD_EMAIL || process.env.SMTP_USER,
      subject: 'New Franchise Inquiry',
      text: `Name: ${name}\nEmail: ${email}\nPhone: ${phone}\nCity: ${city}\nMessage: ${message}`,
      attachments: [
        {
          filename: 'leads.xlsx',
          path: workbookPath
        }
      ]
    };

    await transporter.sendMail(mailOptions);

    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: 'Failed to process lead' });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log(`Server started on port ${port}`));
