export default async function handler(req, res) {

  // =========================================================
  // ONLY POST REQUESTS
  // =========================================================

  if (req.method !== 'POST') {

    return res.status(405).json({
      success: false,
      error: 'POST only'
    });

  }


  // =========================================================
  // GOOGLE APPS SCRIPT URL
  // =========================================================

  const url =
    process.env.GOOGLE_APPS_SCRIPT_URL;


  if (!url) {

    return res.status(503).json({
      success: false,
      error:
        'GOOGLE_APPS_SCRIPT_URL is not configured'
    });

  }


  try {

    const incoming =
      req.body || {};


    // =======================================================
    // BILL HISTORY SEARCH
    // =======================================================

    if (
      incoming.action ===
      'searchBills'
    ) {

      const payload = {

        action:
          'searchBills',

        search:
          incoming.search || {}

      };


      const response =
        await fetch(url, {

          method: 'POST',

          headers: {
            'Content-Type':
              'application/json'
          },

          body:
            JSON.stringify(payload)

        });


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


      return res
        .status(
          response.ok
            ? 200
            : 502
        )
        .json(result);

    }



    // =======================================================
    // TODAY'S BILL OVERVIEW
    // =======================================================

    if (
      incoming.action ===
      'todayBills'
    ) {

      const payload = {

        action:
          'todayBills'

      };


      const response =
        await fetch(url, {

          method: 'POST',

          headers: {
            'Content-Type':
              'application/json'
          },

          body:
            JSON.stringify(payload)

        });


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


      return res
        .status(
          response.ok
            ? 200
            : 502
        )
        .json(result);

    }



    // =======================================================
    // SAVE BILL
    // =======================================================

    let bill;


    /*
     * Frontend can send:
     *
     * {
     *   bill: {...}
     * }
     *
     * OR directly:
     *
     * {
     *   items: [...]
     * }
     */

    if (
      incoming.bill &&
      typeof incoming.bill === 'object' &&
      !Array.isArray(incoming.bill)
    ) {

      bill =
        incoming.bill;

    } else {

      bill =
        incoming;

    }


    // =======================================================
    // VALIDATE BILL
    // =======================================================

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


    // =======================================================
    // ENSURE ITEMS ARRAY
    // =======================================================

    if (
      !Array.isArray(
        bill.items
      )
    ) {

      bill.items = [];

    }


    // =======================================================
    // EMPTY BILL CHECK
    // =======================================================

    if (
      bill.items.length === 0
    ) {

      return res.status(400).json({

        success: false,

        error:
          'Bill contains no items'

      });

    }


    // =======================================================
    // PAYMENT MODE CHECK
    // =======================================================

    const paymentMode =
      bill.paymentMode ||
      bill.payment ||
      '';


    if (!paymentMode) {

      return res.status(400).json({

        success: false,

        error:
          'Payment mode is required'

      });

    }


    // =======================================================
    // SEND BILL TO GOOGLE APPS SCRIPT
    // =======================================================

    const payload = {

      action:
        'saveBill',

      bill:
        bill

    };


    const response =
      await fetch(url, {

        method: 'POST',

        headers: {

          'Content-Type':
            'application/json'

        },

        body:
          JSON.stringify(payload)

      });


    // =======================================================
    // READ RESPONSE
    // =======================================================

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


    // =======================================================
    // RETURN RESULT TO WEBSITE
    // =======================================================

    return res
      .status(
        response.ok
          ? 200
          : 502
      )
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
