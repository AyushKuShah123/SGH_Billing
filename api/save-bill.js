export default async function handler(req, res) {

  if (req.method !== 'POST') {
    return res.status(405).json({
      success: false,
      error: 'POST only'
    });
  }

  const url = process.env.GOOGLE_APPS_SCRIPT_URL;

  if (!url) {
    return res.status(503).json({
      success: false,
      error: 'GOOGLE_APPS_SCRIPT_URL is not configured'
    });
  }

  try {

    /*
     * The frontend normally sends:
     *
     * {
     *   id: "...",
     *   items: [...]
     * }
     *
     * But this also supports:
     *
     * {
     *   bill: {
     *      items: [...]
     *   }
     * }
     *
     * This prevents the "Bill contains no items"
     * problem caused by different payload formats.
     */

    const incoming = req.body || {};

    let bill = incoming;

    if (
      incoming.bill &&
      typeof incoming.bill === 'object' &&
      !Array.isArray(incoming.items)
    ) {
      bill = incoming.bill;
    }


    /*
     * Make absolutely sure items is an array.
     */
    if (!Array.isArray(bill.items)) {
      bill.items = [];
    }


    /*
     * Send the exact format expected by Apps Script.
     */
    const payload = {
      action: 'saveBill',
      bill: bill
    };


    const response = await fetch(url, {

      method: 'POST',

      headers: {
        'Content-Type': 'application/json'
      },

      body: JSON.stringify(payload)

    });


    const text = await response.text();


    /*
     * Return Apps Script's actual response
     * back to the website.
     */
    let result;

    try {
      result = JSON.parse(text);
    } catch (e) {

      return res.status(502).json({
        success: false,
        error: 'Invalid response from Google Apps Script',
        response: text
      });

    }


    return res.status(response.ok ? 200 : 502).json(result);


  } catch (error) {

    console.error(
      'Google Apps Script error:',
      error
    );


    return res.status(502).json({

      success: false,

      error:
        'Google Sheets request failed',

      details:
        error.message

    });

  }

}
