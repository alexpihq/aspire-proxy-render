const express = require('express');
const axios = require('axios');
const app = express();
app.use(express.json());

// Функция для получения access token
async function getAccessToken() {
  try {
    const response = await axios.post('https://api.aspireapp.com/public/v1/login', {
      grant_type: 'client_credentials',
      client_id: process.env.ASPIRE_CLIENT_ID,
      client_secret: process.env.ASPIRE_CLIENT_SECRET
    }, {
      headers: {
        'Content-Type': 'application/json'
      }
    });
    
    return response.data.access_token;
  } catch (error) {
    console.error('❌ Token Error:', error.response?.status, error.message);
    throw error;
  }
}

app.get('/aspire', async (req, res) => {
  const { account_id, start_date } = req.query;

  try {
    // Получаем access token
    const accessToken = await getAccessToken();
    
    const response = await axios.get('https://api.aspireapp.com/public/v1/transactions', {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'User-Agent': 'Mozilla/5.0'
      },
      params: {
        account_id,
        start_date
      }
    });

    res.json(response.data);
  } catch (error) {
    console.error('❌ Aspire API Error:', error.response?.status, error.message);
    res.status(error.response?.status || 500).json({
      error: error.message,
      status: error.response?.status || 500
    });
  }
});

app.get('/ping', (req, res) => {
  res.send('✅ Aspire proxy is up');
});

app.get('/token', async (req, res) => {
  try {
    const accessToken = await getAccessToken();
    res.json({ access_token: accessToken });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Aspire proxy running on port ${PORT}`);
});

app.get("/ip", async (req, res) => {
  try {
    const ipRes = await axios.get("https://api.ipify.org?format=json");
    res.json({ egress_ip: ipRes.data.ip });
  } catch (e) {
    res.status(500).json({ error: "Failed to fetch IP" });
  }
});
