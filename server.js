const express = require("express");
const mongoose = require("mongoose");
const cors = require("cors");

const app = express();

// Middleware
app.use(cors());
app.use(express.json());

// MongoDB connection
mongoose.connect("mongodb://127.0.0.1:27017/constructionDB")
  .then(() => console.log("MongoDB connected"))
  .catch(err => console.log(err));

// Schema
const ContactSchema = new mongoose.Schema({
  name: String,
  phone: String,
  email: String,
  message: String
});

const Contact = mongoose.model("Contact", ContactSchema);

// Route
app.post("/contact", async (req, res) => {
  try {
    const data = new Contact(req.body);
    await data.save();
    res.json({ message: "Enquiry submitted successfully!" });
  } catch (error) {
    res.status(500).json({ message: "Error saving data" });
  }
});

// Start server
app.listen(3000, () => {
  console.log("Server running on http://localhost:3000");
});

// ADMIN – VIEW ALL ENQUIRIES
app.get("/admin/enquiries", async (req, res) => {
  try {
    const enquiries = await Contact.find().sort({ createdAt: -1 });
    res.json(enquiries);
  } catch (error) {
    res.status(500).json({ message: "Error fetching enquiries" });
  }
});

const ExcelJS = require("exceljs");

// EXPORT ENQUIRIES TO EXCEL
app.get("/admin/export", async (req, res) => {
  try {
    const enquiries = await Contact.find().sort({ createdAt: -1 });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Enquiries");

    worksheet.columns = [
      { header: "Name", key: "name", width: 25 },
      { header: "Phone", key: "phone", width: 18 },
      { header: "Email", key: "email", width: 30 },
      { header: "Message", key: "message", width: 40 },
      { header: "Date", key: "createdAt", width: 22 }
    ];

    enquiries.forEach(e => {
      worksheet.addRow({
        name: e.name,
        phone: e.phone,
        email: e.email,
        message: e.message,
        createdAt: new Date(e.createdAt).toLocaleString()
      });
    });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=enquiries.xlsx"
    );

    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    res.status(500).json({ message: "Excel export failed" });
  }
});
