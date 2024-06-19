require('dotenv').config();
const express = require('express');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const axios = require('axios');
const cors = require('cors');

const app = express();

// Define allowed origins
const allowedOrigins = [
    'https://techtribe.powerappsportals.com',
    'http://127.0.0.1:5500',
    'https://datatest.powerappsportals.com'
];

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

// Existing subscription endpoint
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

// New endpoints

// Endpoint to fetch quiz questions
app.get('/quiz', async (req, res) => {
    try {
        const authResult = await client.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"],
        });

        const result = await axios.get(`https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.QUIZ_LIST_ID}/items`, {
            headers: {
                Authorization: `Bearer ${authResult.accessToken}`
            },
            params: {
                expand: 'fields'
            }
        });

        const questions = result.data.value.map(item => ({
            id: item.id,
            title: item.fields.Title,
            optionA: item.fields.field_1,
            optionB: item.fields.field_2,
            optionC: item.fields.field_3,
            optionD: item.fields.field_4,
            correctAnswer: item.fields.field_5
        }));

        res.json(questions);
    } catch (error) {
        console.error('Error fetching quiz questions:', error);
        res.status(500).json({ error: error.message });
    }
});

// Endpoint to submit student details
app.post('/submit-details', async (req, res) => {
    const { fullName, email } = req.body;
    if (!fullName || !email) {
        return res.status(400).json({ error: "Full name and email are required" });
    }

    try {
        const authResult = await client.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"],
        });

        // Check if the email already exists in the StudentExams list
        const checkResponse = await axios.get(`https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.STUDENT_LIST_ID}/items`, {
            headers: {
                Authorization: `Bearer ${authResult.accessToken}`
            },
            params: {
                filter: `fields/Password eq '${email}'`
            }
        });

        if (checkResponse.data.value.length > 0) {
            return res.status(400).json({ error: "Email already exists. You cannot take the quiz twice." });
        }

        const result = await axios.post(`https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.STUDENT_LIST_ID}/items`, {
            fields: {
                Title: fullName,
                Password: email
            }
        }, {
            headers: {
                Authorization: `Bearer ${authResult.accessToken}`,
                'Content-Type': 'application/json'
            }
        });

        res.json(result.data);
    } catch (error) {
        console.error('Error submitting student details:', error);
        res.status(500).json({ error: error.message });
    }
});

// Endpoint to submit quiz answers
app.post('/submit-answers', async (req, res) => {
    const { email, answers } = req.body;
    if (!email || !answers) {
        return res.status(400).json({ error: "Email and answers are required" });
    }

    try {
        const authResult = await client.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"],
        });

        for (const answer of answers) {
            const requestBody = {
                fields: {
                    Title: email,
                    Answers: answer.questionId,
                    SelectedAnswer: answer.selectedAnswer
                }
            };

            //console.log('Submitting answer:', requestBody); // Log the request data for debugging

            const response = await axios.post(`https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.ANSWER_LIST_ID}/items`, requestBody, {
                headers: {
                    Authorization: `Bearer ${authResult.accessToken}`,
                    'Content-Type': 'application/json'
                }
            });

            //console.log('Response from SharePoint:', response.data); // Log the response data for debugging
        }

        res.json({ success: true });
    } catch (error) {
        console.error('Error submitting quiz answers:', error);
        res.status(500).json({ error: error.message });
    }
});

const port = process.env.PORT || 5200; // Ensure the port is set by environment variable with a fallback
app.listen(port, () => console.log(`Server running on port ${port}`));
