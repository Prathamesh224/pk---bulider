const mongoose = require("mongoose");

// Define the schema
const contactSchema = new mongoose.Schema({
  name: { type: String, required: true },
  email: { type: String, required: true },
  message: { type: String, required: true },
  date: { type: Date, default: Date.now }
});

// Export model
module.exports = mongoose.model("Contact", contactSchema);
