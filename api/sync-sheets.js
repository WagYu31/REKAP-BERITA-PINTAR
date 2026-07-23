const axios = require('axios');

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { webhookUrl, mediaName, period, items } = req.body || {};

  if (!webhookUrl) {
    return res.status(400).json({ error: 'webhookUrl wajib diisi' });
  }

  try {
    const response = await axios.post(webhookUrl, { mediaName, period, items }, {
      headers: { 'Content-Type': 'application/json' },
      timeout: 15000,
      maxRedirects: 5
    });

    res.json({ success: true, responseData: response.data });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
