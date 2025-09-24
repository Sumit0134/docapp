const express = require("express");
const multer = require("multer");
const path = require("path");
const cors = require("cors");
const fs = require("fs");
const nodemailer = require("nodemailer");
const ExcelJS = require("exceljs");
const cloudinary = require("cloudinary").v2;
require("dotenv").config();

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.static("public"));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Configure Cloudinary
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET,
});

// Setup Multer (in-memory)
const storage = multer.memoryStorage();
const upload = multer({ storage });

// Nodemailer setup
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
});

// Main form handler
app.post("/submit", upload.single("document"), async (req, res) => {
  const { name, dob, phone, ssn } = req.body;

  if (
    !name.match(/^[A-Za-z ]{3,}$/) || // keep name rule
    !dob || // just check it's not empty
    !phone.match(/^\d{10}$/) || // keep 10-digit phone
    !ssn // just check it's not empty
  ) {
    return res.status(400).send("Invalid form data");
  }

  if (!req.file) {
    return res.status(400).send("No document uploaded");
  }

  try {
    // 1. Upload file to Cloudinary
    const uploadRes = await cloudinary.uploader.upload_stream(
      {
        resource_type: "auto",
        folder: "form-uploads",
      },
      async (error, result) => {
        if (error) {
          console.error("Cloudinary Error:", error);
          return res.status(500).send("File upload failed");
        }

        const fileUrl = result.secure_url;

        // 2. Create Excel file in memory
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet("Submission");

        const timestamp = new Date().toLocaleString("en-IN", {
          timeZone: "Asia/Kolkata",
        });

        sheet.columns = [
          { header: "Name", key: "name", width: 25 },
          { header: "Date of Birth", key: "dob", width: 20 },
          { header: "Phone", key: "phone", width: 15 },
          { header: "SSN", key: "ssn", width: 20 },
          { header: "Document Link", key: "link", width: 25 },
          { header: "Submitted At", key: "timestamp", width: 25 },
        ];

        const row = sheet.addRow({
          name,
          dob,
          phone,
          ssn,
          timestamp,
        });

        row.getCell("link").value = {
          text: "View Document",
          hyperlink: fileUrl,
        };
        row.getCell("link").font = {
          color: { argb: "FF0000FF" },
          underline: true,
        };

        const buffer = await workbook.xlsx.writeBuffer();

        // 3. Send Email with attachment
        await transporter.sendMail({
          from: process.env.EMAIL_USER,
          to: process.env.RECEIVER_EMAIL,
          subject: "New Form Submission",
          text: `New form submission received from ${name}.`,
          attachments: [
            {
              filename: "submission.xlsx",
              content: buffer,
              contentType:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
          ],
        });

        res.status(200).send("Form submitted and email sent!");
      }
    );

    // Write the file to the upload_stream
    uploadRes.end(req.file.buffer);
  } catch (err) {
    console.error("Error:", err);
    res.status(500).send("Something went wrong");
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
});
