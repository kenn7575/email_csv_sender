require('dotenv').config();
const axios = require('axios');
const fs = require('fs');
const { createObjectCsvWriter } = require('csv-writer');
const ExcelJS = require('exceljs');
const nodemailer = require('nodemailer');

async function fetchData() {
  const response = await axios.get(process.env.API_URL);
  console.log("Fetched data:", response.data);
  return response.data;
}

async function writeCsv(data) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("No data to write to CSV.");
  }

  const csvWriter = createObjectCsvWriter({
    path: 'output.csv',
    header: Object.keys(data[0]).map(key => ({ id: key, title: key }))
  });

  await csvWriter.writeRecords(data);
  console.log('CSV file written.');
}

async function writeExcel(data) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('API Data');

  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("No data to write to Excel.");
  }

  worksheet.columns = Object.keys(data[0]).map(key => ({ header: key, key: key }));
  data.forEach(item => worksheet.addRow(item));

  await workbook.xlsx.writeFile('output.xlsx');
  console.log('Excel file written.');
}

async function sendEmail() {
  let transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT),
    secure: process.env.SMTP_SECURE === 'true',
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS
    }
  });

  await transporter.sendMail({
    from: process.env.SMTP_USER,
    to: process.env.RECIPIENT_EMAIL,
    subject: 'CSV and Excel Data Report',
    text: 'Attached are the CSV and Excel files with the latest data report.',
    attachments: [
      { filename: 'output.csv', path: './output.csv' },
      { filename: 'output.xlsx', path: './output.xlsx' }
    ]
  });

  console.log('Email sent with attachments.');
}

(async () => {
  try {
    const data = await fetchData();
    await writeCsv(data.data);
    await writeExcel(data.data);
    await sendEmail();
  } catch (error) {
    console.error('Error:', error);
  }
})();
