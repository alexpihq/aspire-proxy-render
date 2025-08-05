const express = require('express');
const axios = require('axios');
const app = express();
app.use(express.json());

// Кэш для токена
let tokenCache = {
  access_token: null,
  expires_at: null
};

// Функция для получения access token
async function getAccessToken() {
  // Проверяем, есть ли валидный токен в кэше
  if (tokenCache.access_token && tokenCache.expires_at && Date.now() < tokenCache.expires_at) {
    console.log('✅ Using cached token');
    return tokenCache.access_token;
  }

  try {
    console.log('🔄 Getting new access token...');
    const response = await axios.post('https://api.aspireapp.com/public/v1/login', {
      grant_type: 'client_credentials',
      client_id: process.env.ASPIRE_CLIENT_ID,
      client_secret: process.env.ASPIRE_CLIENT_SECRET
    }, {
      headers: {
        'Content-Type': 'application/json'
      }
    });
    
    // Кэшируем токен с учетом времени жизни
    const expiresIn = parseInt(response.data.expires_in) * 1000; // конвертируем в миллисекунды
    tokenCache = {
      access_token: response.data.access_token,
      expires_at: Date.now() + expiresIn - 60000 // вычитаем 1 минуту для безопасности
    };
    
    console.log(`✅ Token cached, expires in ${Math.round(expiresIn/1000)}s`);
    return response.data.access_token;
  } catch (error) {
    console.error('❌ Token Error:', error.response?.status, error.message);
    throw error;
  }
}

app.get('/aspire', async (req, res) => {
  const { account_id, start_date } = req.query;

  console.log('📡 Request received:', { account_id, start_date });

  try {
    // Получаем access token
    const accessToken = await getAccessToken();
    console.log('🔑 Token obtained, length:', accessToken.length);
    
    console.log('🌐 Making request to Aspire API with params:', { account_id, start_date });
    
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

    console.log('✅ Aspire API response status:', response.status);
    console.log('📊 Response data keys:', Object.keys(response.data || {}));
    
    res.json(response.data);
  } catch (error) {
    console.error('❌ Aspire API Error:', error.response?.status, error.message);
    console.error('❌ Error details:', {
      status: error.response?.status,
      statusText: error.response?.statusText,
      data: error.response?.data,
      url: error.config?.url,
      params: error.config?.params
    });
    
    res.status(error.response?.status || 500).json({
      error: error.message,
      status: error.response?.status || 500,
      details: error.response?.data
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
