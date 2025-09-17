const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

// Enable CORS for all origins (adjust for production)
app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'PUT', 'DELETE'],
    allowedHeaders: ['Content-Type', 'Authorization']
}));

app.use(bodyParser.json());

// Serve static files from current directory (adjust path as needed)
app.use(express.static(__dirname));

// In-memory store for saved texts
let store = {};
let nextId = 1;

app.use(express.static(path.join(__dirname, "../src")));

// Root route
app.get('/', (req, res) => {
    res.send('PowerPoint Add-in Server is running');
});

// Serve taskpane.html
app.get('/taskpane.html', (req, res) => {
    res.sendFile(path.join(__dirname, '../src/taskpane.html'));
});

// Create
app.post('/api/texts', (req, res) => {
    try {
        const { text } = req.body;
        if (typeof text !== 'string') {
            return res.status(400).json({ error: 'text required' });
        }
        const id = String(nextId++);
        store[id] = text;
        console.log(`Created text with ID ${id}: ${text.substring(0, 50)}...`);
        res.json({ id, text });
    } catch (error) {
        console.error('Error creating text:', error);
        res.status(500).json({ error: 'Internal server error' });
    }
});

// Read list
app.get('/api/texts', (req, res) => {
    try {
        const items = Object.keys(store).map(id => ({ id, text: store[id] }));
        console.log(`Retrieved ${items.length} texts`);
        res.json(items);
    } catch (error) {
        console.error('Error retrieving texts:', error);
        res.status(500).json({ error: 'Internal server error' });
    }
});

// Server health
app.get('/health', (req, res) => {
    res.json({ health: "OK", timestamp: new Date().toISOString() });
});

// Read single
app.get('/api/texts/:id', (req, res) => {
    try {
        const id = req.params.id;
        if (!(id in store)) {
            return res.status(404).json({ error: 'not found' });
        }
        console.log(`Retrieved text with ID ${id}`);
        res.json({ id, text: store[id] });
    } catch (error) {
        console.error('Error retrieving text:', error);
        res.status(500).json({ error: 'Internal server error' });
    }
});

// Update
app.put('/api/texts/:id', (req, res) => {
    try {
        const id = req.params.id;
        if (!(id in store)) {
            return res.status(404).json({ error: 'not found' });
        }
        const { text } = req.body;
        if (typeof text !== 'string') {
            return res.status(400).json({ error: 'text required' });
        }
        store[id] = text;
        console.log(`Updated text with ID ${id}`);
        res.json({ id, text });
    } catch (error) {
        console.error('Error updating text:', error);
        res.status(500).json({ error: 'Internal server error' });
    }
});

// Delete
app.delete('/api/texts/:id', (req, res) => {
    try {
        const id = req.params.id;
        if (!(id in store)) {
            return res.status(404).json({ error: 'not found' });
        }
        delete store[id];
        console.log(`Deleted text with ID ${id}`);
        res.json({ ok: true });
    } catch (error) {
        console.error('Error deleting text:', error);
        res.status(500).json({ error: 'Internal server error' });
    }
});

// Error handling middleware
app.use((error, req, res, next) => {
    console.error('Unhandled error:', error);
    res.status(500).json({ error: 'Internal server error' });
});

app.listen(PORT, () => {
    console.log(`Server listening on http://localhost:${PORT}`);
    console.log(`Ngrok URL should be accessible at your ngrok domain`);
});