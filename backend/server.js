const express = require('express');
const cors = require('cors');
const path = require('path');
require('dotenv').config();

const documentRoutes = require('./routes/documentRoutes');

const app = express();

app.use(cors());
app.use(express.json());

// Static files serve karne ke liye
app.use(express.static(path.join(__dirname, 'public')));
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

// API Routes
app.use('/api/documents', documentRoutes);

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});