export default async function handler(req, res) {
  // Only allow POST requests
  if (req.method !== 'POST') {
    return res.status(405).json({
      success: false,
      error: 'POST only'
    });
  }

  // Apps Script Web App URL
  const url = process.env.GOOGLE_APPS_SCRIPT_URL;

  if (!url) {
    return res.status(503).json({
      success: false,
      error: 'Google Apps Script URL is not configured'
    });
  }

  try {
    // Get bill sent from the website
    const bill = req.body;

    // Send the correct structure to Google Apps Script
    const response = await fetch(url, {
      method: 'POST',

      headers: {
        'Content-Type': 'application/json'
      },

      body: JSON.stringify({
        action: 'saveBill',
        bill: bill
      })
    });

    const text = await response.text();

    // Return Apps Script response to the website
    return res
      .status(response.ok ? 200 : 502)
      .send(text);

  } catch (error) {

    console.error('Google Apps Script error:', error);

    return res.status(502).json({
      success: false,
      error: 'Google Sheets request failed',
      details: error.message
    });
  }
}
