const express = require('express');
const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { Octokit } = require('@octokit/rest');

const octokit = new Octokit({ auth: process.env.GITHUB_TOKEN });
const githubRepo = process.env.GITHUB_REPO || '';

const app = express();
app.use(express.static(__dirname));
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
    }
  } else {
    worksheet = workbook.addWorksheet('Leads');
    await workbook.xlsx.writeFile(workbookPath);
  }
  worksheet.columns = [
    { header: 'Name', key: 'name', width: 20 },
    { header: 'Email', key: 'email', width: 30 },
    { header: 'Phone', key: 'phone', width: 15 },
    { header: 'State', key: 'state', width: 20 },
    { header: 'City', key: 'city', width: 20 },
    { header: 'Source', key: 'source', width: 20 },
    { header: 'Other Source', key: 'otherSource', width: 20 },
    { header: 'Message', key: 'message', width: 50 },
    { header: 'Date', key: 'date', width: 20 }
  ];
}

async function startServer() {
  await initWorkbook();
  const port = process.env.PORT || 3000;
  app.listen(port, () => console.log(`Server started on port ${port}`));
}

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
  const { name, email, phone, city, state, source, otherSource, message } = req.body;
  try {
    worksheet.addRow({
      name,
      email,
      phone,
      state,
      city,
      source,
      otherSource,
      message,
      date: new Date().toISOString()
    });
    await workbook.xlsx.writeFile(workbookPath);

    const mailOptions = {
      from: process.env.SMTP_FROM || process.env.SMTP_USER,
      to: process.env.LEAD_EMAIL || 'digimarketing@safc.ltd',
      subject: 'New Franchise Inquiry',
      text: `Name: ${name}\nEmail: ${email}\nPhone: ${phone}\nState: ${state}\nCity: ${city}\nSource: ${source}${otherSource ? ` (${otherSource})` : ''}\nMessage: ${message}`,
      attachments: [
        {
          filename: 'leads.xlsx',
          path: workbookPath
        }
      ]
    };

    if (process.env.SMTP_USER && process.env.SMTP_PASS && process.env.SMTP_HOST) {
      await transporter.sendMail(mailOptions);
    } else {
      console.log('SMTP credentials not configured. Skipping email sending.');
    }

    if (process.env.GITHUB_TOKEN && githubRepo) {
      const [owner, repo] = githubRepo.split('/');
      try {
        await octokit.issues.create({
          owner,
          repo,
          title: `New lead from ${name}`,
          body: `Name: ${name}\nEmail: ${email}\nPhone: ${phone}\nState: ${state}\nCity: ${city}\nSource: ${source}${otherSource ? ` (${otherSource})` : ''}\nMessage: ${message}`
        });
      } catch (ghErr) {
        console.error('Failed to create GitHub issue', ghErr);
      }
    }

    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: 'Failed to process lead' });
  }
});
startServer();
