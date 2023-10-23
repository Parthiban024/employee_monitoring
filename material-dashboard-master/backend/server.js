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
const dbName = 'Employee_monitoring';
app.post('/api/upload', upload.single('excelFile'), async (req, res) => {
  try {
    const buffer = req.file.buffer;
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const excelData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Compare Excel headers with table headers
    const excelHeaders = excelData[0];
    const tableHeaders = ["EmpId", "Name", "ProjectName", "Department", 'Date', 'Status', 'ProductionStatus']; // Your table headers

    if (!compareHeaders(excelHeaders, tableHeaders)) {
      return res.status(400).json({ error: "Excel headers do not match table headers." });
    }

    // Connect to MongoDB Atlas
    const client = new MongoClient(mongoURL, { useNewUrlParser: true, useUnifiedTopology: true });
    await client.connect();

    // Insert data into MongoDB (with filtering to exclude rows with mostly null values)
    const db = client.db(dbName);
    const collection = db.collection('empmonitor');
    const dataToInsert = excelData.slice(1).filter((row) => {
      // Customize this condition based on your requirements
      // This example assumes that a row should have at least one non-null field
      return Object.values(row).some((value) => value !== null && value !== undefined);
    }).map((row) => {
      const obj = {};
      for (let i = 0; i < tableHeaders.length; i++) {
        obj[tableHeaders[i]] = row[i];
      }
      return obj;
    });

    // Check for duplicates based on EmpId and Name
    const duplicates = await findDuplicates(collection, dataToInsert);

    if (duplicates.length > 0) {
      return res.status(400).json({ error: "Duplicate data detected. Cannot upload duplicates." });
    }

    // Insert the data into MongoDB
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

// Function to find duplicates based on EmpId and Name
async function findDuplicates(collection, dataToInsert) {
  const empIds = dataToInsert.map((item) => item.EmpId);
  const names = dataToInsert.map((item) => item.Name);

  const duplicates = await collection.find({
    $or: [
      { EmpId: { $in: empIds } },
      { Name: { $in: names } }
    ]
  }).toArray();

  return duplicates;
}
// Add a route to fetch data from MongoDB and send it as JSON
app.get('/api/upload', async (req, res) => {
  try {
    const client = new MongoClient(mongoURL, { useUnifiedTopology: true });
    await client.connect();

    const db = client.db(dbName);
    const collection = db.collection('empmonitor'); // Change 'empmon' to your collection name

    const data = await collection.find({}).toArray();

    client.close();

    res.json(data);
  } catch (error) {
    res.status(500).json({ error: "An error occurred." });
  }
});

// Add a route to delete data by _id
app.delete('/api/delete/:id', async (req, res) => {
  try {
    const { id } = req.params;

    const client = new MongoClient(mongoURL, { useUnifiedTopology: true });
    await client.connect();

    const db = client.db(dbName);
    const collection = db.collection('empmonitor'); // Change to your collection name

    // Perform the delete operation
    const result = await collection.deleteOne({ _id: new mongodb.ObjectId(id) }); // Updated from mongodb.ObjectID to mongodb.ObjectId

    client.close();

    if (result.deletedCount === 1) {
      res.status(200).json({ success: "Data deleted successfully." });
    } else {
      res.status(404).json({ error: "Data not found." });
    }
  } catch (error) {
    console.error(error);
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