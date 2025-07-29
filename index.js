const express = require('express');
const axios = require('axios');
const app = express();
app.use(express.json());

app.get('/aspire', async (req, res) => {
  const { account_id, start_date } = req.query;

  try {
    const response = await axios.get('https://api.aspireapp.com/public/v1/transactions', {
      headers: {
        'x-api-key': process.env.ASPIRE_API_KEY,
        'x-client-id': process.env.ASPIRE_CLIENT_ID,
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
