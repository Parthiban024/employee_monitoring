const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const mongodb = require('mongodb');
const cors = require('cors');
const MongoClient = mongodb.MongoClient;

const app = express();
app.use(cors());
const port = 8080;

// Configure multer for file uploads
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// MongoDB Atlas connection URL (replace with your URL)
const mongoURL = 'mongodb+srv://ParthiGMR:Parthiban7548@parthibangmr.1quwer2.mongodb.net';

// Specify the database name separately
const dbName = 'IT-tool';

app.post('/api/upload', upload.single('excelFile'), async (req, res) => {
  try {
    const buffer = req.file.buffer;
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const excelData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Compare Excel headers with table headers
    const excelHeaders = excelData[0];
    const tableHeaders = ["Author", "Function", "Status", "Employed"]; // Your table headers

    if (!compareHeaders(excelHeaders, tableHeaders)) {
      return res.status(400).json({ error: "Excel headers do not match table headers." });
    }

    // Connect to MongoDB Atlas
    const client = new MongoClient(mongoURL, { useNewUrlParser: true, useUnifiedTopology: true });
    await client.connect();

    // Insert data into MongoDB (you need to implement MongoDB logic here)
    const db = client.db(dbName); // Specify the database name here
    const collection = db.collection('empmon');
    const dataToInsert = excelData.slice(1).map((row) => {
      const obj = {};
      for (let i = 0; i < tableHeaders.length; i++) {
        obj[tableHeaders[i]] = row[i];
      }
      return obj;
    });

    await collection.insertMany(dataToInsert);

    // Close the MongoDB connection
    client.close();

    // Send a success message
    res.status(200).json({ success: "Data uploaded successfully. MongoDB connected successfully." });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "An error occurred." });
  }
});
// Add a route to fetch data from MongoDB and send it as JSON
app.get('/api/upload', async (req, res) => {
  try {
    const client = new MongoClient(mongoURL, { useUnifiedTopology: true });
    await client.connect();

    const db = client.db(dbName);
    const collection = db.collection('empmon'); // Change 'empmon' to your collection name

    const data = await collection.find({}).toArray();

    client.close();

    res.json(data);
  } catch (error) {
    res.status(500).json({ error: "An error occurred." });
  }
});


function compareHeaders(headers1, headers2) {
  if (headers1.length !== headers2.length) {
    return false;
  }

  for (let i = 0; i < headers1.length; i++) {
    if (headers1[i] !== headers2[i]) {
      return false;
    }
  }

  return true;
}

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
