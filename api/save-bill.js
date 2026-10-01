export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({error:'POST only'});
  const url = process.env.GOOGLE_APPS_SCRIPT_URL;
  const secret = process.env.GOOGLE_SHEETS_SECRET;
  if (!url || !secret) return res.status(503).json({error:'Google Sheets integration is not configured'});
  try {
    const response = await fetch(url, {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({secret,bill:req.body})});
    const text = await response.text();
    return res.status(response.ok ? 200 : 502).send(text);
  } catch (e) { return res.status(502).json({error:'Google Sheets request failed'}); }
}
