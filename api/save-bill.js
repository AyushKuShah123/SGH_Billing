export default async function handler(req, res) {

  // Only POST requests are allowed
  if (req.method !== 'POST') {
    return res.status(405).json({
      success: false,
      error: 'POST only'
    });
  }


  // Google Apps Script Web App URL
  const url = process.env.GOOGLE_APPS_SCRIPT_URL;


  // Check environment variable
  if (!url) {
    return res.status(503).json({
      success: false,
      error: 'GOOGLE_APPS_SCRIPT_URL is not configured'
    });
  }


  try {

    const incoming = req.body || {};


    /*
     * =========================================================
     * BILL HISTORY SEARCH
     * =========================================================
     *
     * Frontend sends:
     *
     * {
     *   action: "searchBills",
     *   search: {
     *      date: "...",
     *      fromDate: "...",
     *      toDate: "...",
     *      customer: "...",
     *      billNo: "...",
     *      paymentMode: "..."
     *   }
     * }
     *
     * Forward it directly to Google Apps Script.
     */

    if (incoming.action === 'searchBills') {

      const payload = {
        action: 'searchBills',
        search: incoming.search || {}
      };


      const response = await fetch(url, {

        method: 'POST',

        headers: {
          'Content-Type': 'application/json'
        },

        body: JSON.stringify(payload)

      });


      const text = await response.text();


      let result;

      try {

        result = JSON.parse(text);

      } catch (e) {

        return res.status(502).json({

          success: false,

          error:
            'Invalid response from Google Apps Script',

          response:
            text

        });

      }


      return res
        .status(response.ok ? 200 : 502)
        .json(result);

    }



    /*
     * =========================================================
     * SAVE BILL
     * =========================================================
     *
     * The frontend may send either:
     *
     * 1. Direct bill:
     *
     * {
     *   id: "...",
     *   items: [...]
     * }
     *
     * OR
     *
     * 2. Wrapped bill:
     *
     * {
     *   bill: {
     *      id: "...",
     *      items: [...]
     *   }
     * }
     */


    let bill;


    if (
      incoming.bill &&
      typeof incoming.bill === 'object' &&
      !Array.isArray(incoming.bill)
    ) {

      bill = incoming.bill;

    } else {

      bill = incoming;

    }


    /*
     * Make sure bill is an object.
     */

    if (
      !bill ||
      typeof bill !== 'object' ||
      Array.isArray(bill)
    ) {

      return res.status(400).json({

        success: false,

        error:
          'Invalid bill data'

      });

    }


    /*
     * Make sure items exists.
     */

    if (!Array.isArray(bill.items)) {

      bill.items = [];

    }


    /*
     * Do not allow an empty bill to be sent.
     *
     * This gives a clear error instead of sending
     * an invalid bill to Google Apps Script.
     */

    if (bill.items.length === 0) {

      return res.status(400).json({

        success: false,

        error:
          'Bill contains no items'

      });

    }


    /*
     * =========================================================
     * SEND BILL TO GOOGLE APPS SCRIPT
     * =========================================================
     */

    const payload = {

      action: 'saveBill',

      bill: bill

    };


    const response = await fetch(url, {

      method: 'POST',

      headers: {

        'Content-Type':
          'application/json'

      },

      body:
        JSON.stringify(payload)

    });


    /*
     * Read Google Apps Script response.
     */

    const text =
      await response.text();


    let result;


    try {

      result =
        JSON.parse(text);

    } catch (e) {

      return res.status(502).json({

        success: false,

        error:
          'Invalid response from Google Apps Script',

        response:
          text

      });

    }


    /*
     * Return the actual Apps Script result
     * back to index.html.
     */

    return res
      .status(response.ok ? 200 : 502)
      .json(result);


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
