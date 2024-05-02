require('dotenv').config();
const express = require('express');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const axios = require('axios');
const cors = require('cors');

const app = express();

// Define allowed origins
const allowedOrigins = ['https://techtribe.powerappsportals.com', 'http://127.0.0.1:5500'];

// Configure CORS to accept requests from multiple origins
app.use(cors({
    origin: function (origin, callback) {
        // allow requests with no origin (like mobile apps or curl requests)
        if (!origin) return callback(null, true);
        if (allowedOrigins.indexOf(origin) === -1) {
            var msg = 'The CORS policy for this site does not allow access from the specified Origin.';
            return callback(new Error(msg), false);
        }
        return callback(null, true);
    }
}));

app.use(express.json());

const config = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET,
    }
};

const client = new ConfidentialClientApplication(config);

app.post('/subscribe', async (req, res) => {
    const email = req.body.email;
    if (!email) {
        return res.status(400).json({ error: "Email is required" });
    }

    try {
        const authResult = await client.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"],
        });

        const result = await axios.post(`https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.LIST_ID}/items`, {
            fields: { Title: email }
        }, {
            headers: {
                Authorization: `Bearer ${authResult.accessToken}`,
                'Content-Type': 'application/json'
            }
        });

        res.json(result.data);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Add a simple health check endpoint
app.get('/health', (req, res) => {
    res.status(200).send('OK');
});

const port = process.env.PORT || 5000; // Ensure the port is set by environment variable with a fallback
app.listen(port, () => console.log(`Server running on port ${port}`));
