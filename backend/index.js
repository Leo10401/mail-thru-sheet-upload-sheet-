const express = require("express")
const multer = require("multer")
const XLSX = require("xlsx")
const cors = require("cors")
require("dotenv").config()
const nodemailer = require("nodemailer")
const path = require("path")
const fs = require("fs")
const mongoose = require("mongoose")
const SendedMail = require("./models/sendedmailModel")

const app = express()
const PORT = process.env.PORT || 3000

// MongoDB connection
mongoose.connect(process.env.MONGODB_URI || "mongodb://localhost:27017/sheetmail")
  .then(() => console.log("Connected to MongoDB"))
  .catch((err) => console.error("MongoDB connection error:", err))

// Middleware
app.use(cors())
app.use(express.json({ limit: "50mb" }))
app.use(express.urlencoded({ extended: true, limit: "50mb" }))

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = "uploads/"
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true })
    }
    cb(null, uploadDir)
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + "-" + Math.round(Math.random() * 1e9)
    cb(null, file.fieldname + "-" + uniqueSuffix + path.extname(file.originalname))
  },
})

const upload = multer({
  storage: storage,
  fileFilter: (req, file, cb) => {
    const allowedTypes = [".xlsx", ".xls", ".csv"]
    const fileExt = path.extname(file.originalname).toLowerCase()
    if (allowedTypes.includes(fileExt)) {
      cb(null, true)
    } else {
      cb(new Error("Only Excel files (.xlsx, .xls) and CSV files are allowed"))
    }
  },
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB limit
  },
})

// SMTP Transporter setup
const smtpTransporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: Number.parseInt(process.env.SMTP_PORT, 10),
  secure: process.env.SMTP_SECURE === "true",
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASSWORD,
  },
})

// Helper function to parse Excel/CSV files
function parseSpreadsheetFile(filePath) {
  try {
    const workbook = XLSX.readFile(filePath)
    const sheetNames = workbook.SheetNames
    const sheets = {}

    sheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName]
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })
      sheets[sheetName] = {
        rawData: jsonData,
        data: convertToObjects(jsonData),
        rowCount: jsonData.length,
        columnCount: jsonData[0]?.length || 0,
      }
    })

    return {
      sheets,
      sheetNames,
      defaultSheet: sheetNames[0],
    }
  } catch (error) {
    throw new Error(`Failed to parse spreadsheet: ${error.message}`)
  }
}

// Helper function to parse data from URL
async function parseSpreadsheetFromUrl(url) {
  try {
    const response = await fetch(url)
    if (!response.ok) {
      throw new Error(`Failed to fetch file from URL: ${response.statusText}`)
    }

    const buffer = await response.arrayBuffer()
    const workbook = XLSX.read(buffer)
    const sheetNames = workbook.SheetNames
    const sheets = {}

    sheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName]
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })
      sheets[sheetName] = {
        rawData: jsonData,
        data: convertToObjects(jsonData),
        rowCount: jsonData.length,
        columnCount: jsonData[0]?.length || 0,
      }
    })

    return {
      sheets,
      sheetNames,
      defaultSheet: sheetNames[0],
    }
  } catch (error) {
    throw new Error(`Failed to parse spreadsheet from URL: ${error.message}`)
  }
}

// Helper function to convert range to object array
function convertToObjects(values, headers = null) {
  if (!values || values.length === 0) return [];

  // If headers are provided, use them directly
  if (headers) {
    const dataRows = values;
    return dataRows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] || '';
      });
      return obj;
    });
  }

  // If no headers provided, check if first row contains column names (A, B, C, etc.)
  const firstRow = values[0];
  const isColumnNames = firstRow.every(cell => 
    typeof cell === 'string' && /^[A-Z]+$/.test(cell)
  );

  if (isColumnNames) {
    // Use second row as headers
    const headerRow = values[1];
    const dataRows = values.slice(2);
    return dataRows.map(row => {
      const obj = {};
      headerRow.forEach((header, index) => {
        obj[header] = row[index] || '';
      });
      return obj;
    });
  } else {
    // Use first row as headers (original behavior)
    const headerRow = values[0];
    const dataRows = values.slice(1);
    return dataRows.map(row => {
      const obj = {};
      headerRow.forEach((header, index) => {
        obj[header] = row[index] || '';
      });
      return obj;
    });
  }
}

// Store uploaded file data in memory (in production, use a database)
const uploadedFiles = new Map()

// Routes

// Health check
app.get("/health", (req, res) => {
  res.json({ status: "OK", message: "XLSX Upload Backend is running" })
})

// Upload Excel file
app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" })
    }

    const fileId = Date.now().toString()
    const parsedData = parseSpreadsheetFile(req.file.path)

    // Store file data
    uploadedFiles.set(fileId, {
      ...parsedData,
      originalName: req.file.originalname,
      uploadedAt: new Date(),
      filePath: req.file.path,
    })

    res.json({
      fileId,
      fileName: req.file.originalname,
      sheets: parsedData.sheetNames,
      message: "File uploaded and parsed successfully",
    })
  } catch (error) {
    console.error("Upload error:", error)
    res.status(500).json({
      error: "Failed to process uploaded file",
      message: error.message,
    })
  }
})

// Upload from URL
app.post("/upload-url", async (req, res) => {
  try {
    const { url } = req.body
    if (!url) {
      return res.status(400).json({ error: "URL is required" })
    }

    const fileId = Date.now().toString()
    const parsedData = await parseSpreadsheetFromUrl(url)

    // Store file data
    uploadedFiles.set(fileId, {
      ...parsedData,
      originalName: `File from ${url}`,
      uploadedAt: new Date(),
      sourceUrl: url,
    })

    res.json({
      fileId,
      fileName: `File from ${url}`,
      sheets: parsedData.sheetNames,
      message: "File loaded from URL and parsed successfully",
    })
  } catch (error) {
    console.error("URL upload error:", error)
    res.status(500).json({
      error: "Failed to process file from URL",
      message: error.message,
    })
  }
})

// Get file data
app.get(["/file/:fileId", "/file/:fileId/:sheetName"], (req, res) => {
  try {
    const { fileId, sheetName } = req.params
    const { format = "objects" } = req.query

    const fileData = uploadedFiles.get(fileId)
    if (!fileData) {
      return res.status(404).json({ error: "File not found" })
    }

    const targetSheet = sheetName || fileData.defaultSheet
    const sheetData = fileData.sheets[targetSheet]

    if (!sheetData) {
      return res.status(404).json({ error: "Sheet not found" })
    }

    const result = {
      rawData: sheetData.rawData,
      data: format === "objects" ? sheetData.data : sheetData.rawData,
      rowCount: sheetData.rowCount,
      columnCount: sheetData.columnCount,
      fileName: fileData.originalName,
      sheetName: targetSheet,
    }

    res.json(result)
  } catch (error) {
    console.error("Error fetching file data:", error)
    res.status(500).json({
      error: "Failed to fetch file data",
      message: error.message,
    })
  }
})

// Get contacts from uploaded file
app.get(["/contacts/:fileId", "/contacts/:fileId/:sheetName"], (req, res) => {
  try {
    const { fileId, sheetName } = req.params
    const { nameColumn = "name", emailColumn = "email", validateEmails = "true" } = req.query

    const fileData = uploadedFiles.get(fileId)
    if (!fileData) {
      return res.status(404).json({ error: "File not found" })
    }

    const targetSheet = sheetName || fileData.defaultSheet
    const sheetData = fileData.sheets[targetSheet]

    if (!sheetData) {
      return res.status(404).json({ error: "Sheet not found" })
    }

    const objects = sheetData.data

    // Email validation regex
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/

    // Extract contacts with name and email
    const contacts = objects
      .map((row, index) => {
        // Try to find name and email columns with flexible matching
        const name =
          row[nameColumn] ||
          row["Name"] ||
          row["name"] ||
          row["Full Name"] ||
          row["full_name"] ||
          row[Object.keys(row).find((key) => key.toLowerCase().includes("name"))] ||
          ""

        const email =
          row[emailColumn] ||
          row["Email"] ||
          row["email"] ||
          row["Email Address"] ||
          row["email_address"] ||
          row[
            Object.keys(row).find((key) => key.toLowerCase().includes("email") || key.toLowerCase().includes("mail"))
          ] ||
          ""

        const certificate =
          row["certificate"] ||
          row["Certificate"] ||
          row["certificate_link"] ||
          row[Object.keys(row).find((key) => key.toLowerCase().includes("certificate"))] ||
          ""

        const isValidEmail = validateEmails === "true" ? emailRegex.test(email) : true

        return {
          id: index + 1,
          name: name.toString().trim(),
          email: email.toString().trim().toLowerCase(),
          certificate: certificate.toString().trim(),
          isValidEmail,
          originalRow: index + 2,
        }
      })
      .filter((contact) => contact.name || contact.email)

    // Statistics
    const validEmails = contacts.filter((c) => c.isValidEmail).length
    const invalidEmails = contacts.filter((c) => c.email && !c.isValidEmail).length
    const contactsWithBoth = contacts.filter((c) => c.name && c.email && c.isValidEmail).length
    const contactsWithCertificates = contacts.filter((c) => c.certificate).length

    res.json({
      contacts,
      totalContacts: contacts.length,
      validEmails,
      invalidEmails,
      contactsWithBoth,
      contactsWithCertificates,
      statistics: {
        hasName: contacts.filter((c) => c.name).length,
        hasEmail: contacts.filter((c) => c.email).length,
        hasBoth: contactsWithBoth,
        hasCertificate: contactsWithCertificates,
        emptyRows: objects.length - contacts.length,
      },
    })
  } catch (error) {
    console.error("Error fetching contacts:", error)
    res.status(500).json({
      error: "Failed to fetch contacts",
      message: error.message,
    })
  }
})

// Get emails from uploaded file
app.get(["/emails/:fileId", "/emails/:fileId/:sheetName"], (req, res) => {
  try {
    const { fileId, sheetName } = req.params
    const { emailColumn = "email", unique = "true" } = req.query

    const fileData = uploadedFiles.get(fileId)
    if (!fileData) {
      return res.status(404).json({ error: "File not found" })
    }

    const targetSheet = sheetName || fileData.defaultSheet
    const sheetData = fileData.sheets[targetSheet]

    if (!sheetData) {
      return res.status(404).json({ error: "Sheet not found" })
    }

    const objects = sheetData.data
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/

    // Extract emails
    let emails = objects
      .map((row) => {
        const email =
          row[emailColumn] ||
          row["Email"] ||
          row["email"] ||
          row["Email Address"] ||
          row["email_address"] ||
          row[
            Object.keys(row).find((key) => key.toLowerCase().includes("email") || key.toLowerCase().includes("mail"))
          ] ||
          ""

        return email.toString().trim().toLowerCase()
      })
      .filter((email) => email && emailRegex.test(email))

    // Remove duplicates if requested
    if (unique === "true") {
      emails = [...new Set(emails)]
    }

    res.json({
      emails,
      totalEmails: emails.length,
      uniqueEmails: unique === "true" ? emails.length : [...new Set(emails)].length,
    })
  } catch (error) {
    console.error("Error fetching emails:", error)
    res.status(500).json({
      error: "Failed to fetch emails",
      message: error.message,
    })
  }
})

// Search contacts in uploaded file
app.get("/contacts/:fileId/:sheetName/search", (req, res) => {
  try {
    const { fileId, sheetName } = req.params
    const { query, type = "both" } = req.query

    if (!query) {
      return res.status(400).json({
        error: "Invalid request",
        message: "query parameter is required",
      })
    }

    const fileData = uploadedFiles.get(fileId)
    if (!fileData) {
      return res.status(404).json({ error: "File not found" })
    }

    const sheetData = fileData.sheets[sheetName]
    if (!sheetData) {
      return res.status(404).json({ error: "Sheet not found" })
    }

    const objects = sheetData.data
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/
    const searchQuery = query.toLowerCase()

    const results = objects
      .map((row, index) => {
        const name =
          row["name"] ||
          row["Name"] ||
          row["Full Name"] ||
          row[Object.keys(row).find((key) => key.toLowerCase().includes("name"))] ||
          ""
        const email =
          row["email"] ||
          row["Email"] ||
          row["Email Address"] ||
          row[
            Object.keys(row).find((key) => key.toLowerCase().includes("email") || key.toLowerCase().includes("mail"))
          ] ||
          ""

        return {
          id: index + 1,
          name: name.toString().trim(),
          email: email.toString().trim().toLowerCase(),
          isValidEmail: emailRegex.test(email),
          originalRow: index + 2,
        }
      })
      .filter((contact) => {
        const nameMatch = contact.name.toLowerCase().includes(searchQuery)
        const emailMatch = contact.email.toLowerCase().includes(searchQuery)

        switch (type) {
          case "name":
            return nameMatch
          case "email":
            return emailMatch
          case "both":
          default:
            return nameMatch || emailMatch
        }
      })

    res.json({
      results,
      totalFound: results.length,
      searchQuery: query,
      searchType: type,
    })
  } catch (error) {
    console.error("Error searching contacts:", error)
    res.status(500).json({
      error: "Failed to search contacts",
      message: error.message,
    })
  }
})

// Send emails to contacts in uploaded file
app.post(["/send-emails/:fileId", "/send-emails/:fileId/:sheetName"], async (req, res) => {
  try {
    const { fileId, sheetName } = req.params
    const { subject, body, templateType } = req.body

    if (!subject || !body) {
      return res.status(400).json({ error: "Subject and body are required" })
    }

    const fileData = uploadedFiles.get(fileId)
    if (!fileData) {
      return res.status(404).json({ error: "File not found" })
    }

    const targetSheet = sheetName || fileData.defaultSheet
    const sheetData = fileData.sheets[targetSheet]

    if (!sheetData) {
      return res.status(404).json({ error: "Sheet not found" })
    }

    const objects = sheetData.data
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/

    // Get contacts with name, email, and certificate link
    const contacts = objects
      .map((row) => {
        const name =
          row["name"] ||
          row["Name"] ||
          row["Full Name"] ||
          row[Object.keys(row).find((key) => key.toLowerCase().includes("name"))] ||
          ""
        const email =
          row["email"] ||
          row["Email"] ||
          row["Email Address"] ||
          row[
            Object.keys(row).find((key) => key.toLowerCase().includes("email") || key.toLowerCase().includes("mail"))
          ] ||
          ""
        const certificateLink =
          row["certificate"] ||
          row["Certificate"] ||
          row[Object.keys(row).find((key) => key.toLowerCase().includes("certificate"))] ||
          ""

        return {
          name: name.toString().trim(),
          email: email.toString().trim().toLowerCase(),
          certificateLink: certificateLink.toString().trim(),
        }
      })
      .filter((contact) => contact.email && emailRegex.test(contact.email) && contact.certificateLink)

    if (contacts.length === 0) {
      return res.status(400).json({ error: "No valid contacts with certificate links found" })
    }

    // Send emails
    let sent = 0,
      failed = 0,
      errors = []
    for (const contact of contacts) {
      try {
        let personalizedBody = body

        // If this is a certificate template, replace the name placeholder and the certificate link
        if (templateType === "certificate") {
          personalizedBody = personalizedBody.replace(/Congratulations!/g, `Congratulations ${contact.name}!`)

          personalizedBody = personalizedBody.replace(
            /<a href="#"([^>]*)>\s*Get Your Certificate\s*<\/a>/g,
            `<a href="${contact.certificateLink}"$1> Get your certificate </a>`,
          )
        }

        await smtpTransporter.sendMail({
          from: process.env.SMTP_USER,
          to: contact.email,
          subject,
          html: personalizedBody,
        })

        // Save email log to MongoDB
        const mailLog = new SendedMail({
          recipientEmail: contact.email,
          recipientName: contact.name,
          certificateDetails: contact.certificateLink,
          sheetName: targetSheet,
          fileName: fileData.originalName
        })
        await mailLog.save()

        sent++
      } catch (err) {
        failed++
        errors.push({ email: contact.email, error: err.message })
      }
    }

    res.json({
      message: `Emails sent: ${sent}, failed: ${failed}`,
      sent,
      failed,
      errors,
    })
  } catch (error) {
    console.error("Error sending emails:", error)
    res.status(500).json({ error: "Failed to send emails", message: error.message })
  }
})

// Get list of uploaded files
app.get("/files", (req, res) => {
  const files = Array.from(uploadedFiles.entries()).map(([id, data]) => ({
    id,
    name: data.originalName,
    sheets: data.sheetNames,
    uploadedAt: data.uploadedAt,
  }))

  res.json({ files })
})

// Delete uploaded file
app.delete("/file/:fileId", (req, res) => {
  const { fileId } = req.params
  const fileData = uploadedFiles.get(fileId)

  if (!fileData) {
    return res.status(404).json({ error: "File not found" })
  }

  // Delete physical file if it exists
  if (fileData.filePath && fs.existsSync(fileData.filePath)) {
    fs.unlinkSync(fileData.filePath)
  }

  uploadedFiles.delete(fileId)
  res.json({ message: "File deleted successfully" })
})

// Get email sending logs with filters
app.get("/email-logs", async (req, res) => {
  try {
    const { 
      timeFilter,
      sheetName, 
      fileName,
      search 
    } = req.query;

    // Build filter object
    const filter = {};

    // Time-based filter
    if (timeFilter) {
      const now = new Date();
      switch (timeFilter) {
        case 'lastHour':
          filter.sentAt = { $gte: new Date(now - 60 * 60 * 1000) };
          break;
        case 'last24Hours':
          filter.sentAt = { $gte: new Date(now - 24 * 60 * 60 * 1000) };
          break;
        case 'lastWeek':
          filter.sentAt = { $gte: new Date(now - 7 * 24 * 60 * 60 * 1000) };
          break;
        case 'lastMonth':
          filter.sentAt = { $gte: new Date(now - 30 * 24 * 60 * 60 * 1000) };
          break;
        case 'last3Months':
          filter.sentAt = { $gte: new Date(now - 90 * 24 * 60 * 60 * 1000) };
          break;
        case 'lastYear':
          filter.sentAt = { $gte: new Date(now - 365 * 24 * 60 * 60 * 1000) };
          break;
      }
    }

    // Sheet name filter
    if (sheetName) {
      filter.sheetName = { $regex: sheetName, $options: 'i' };
    }

    // File name filter
    if (fileName) {
      filter.fileName = { $regex: fileName, $options: 'i' };
    }

    // Search in recipient name or email
    if (search) {
      filter.$or = [
        { recipientName: { $regex: search, $options: 'i' } },
        { recipientEmail: { $regex: search, $options: 'i' } }
      ];
    }

    const logs = await SendedMail.find(filter).sort({ sentAt: -1 });
    res.json(logs);
  } catch (error) {
    console.error("Error fetching email logs:", error);
    res.status(500).json({ error: "Failed to fetch email logs", message: error.message });
  }
});

// Error handling middleware
app.use((error, req, res, next) => {
  console.error("Unhandled error:", error)
  res.status(500).json({
    error: "Internal server error",
    message: error.message,
  })
})

// 404 handler
app.use((req, res) => {
  res.status(404).json({
    error: "Route not found",
    message: `Route ${req.method} ${req.path} not found`,
  })
})

app.listen(PORT, () => {
  console.log(`XLSX Upload Backend running on port ${PORT}`)
  console.log(`Health check: http://localhost:${PORT}/health`)
})

module.exports = app
