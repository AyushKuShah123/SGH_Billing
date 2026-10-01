const SHEET_ID = '1g-StRRD-sm-UWpyQhqTVodKEdmbUrC4vspKxKQoDiYs';


/* =========================================================
   GET
   ========================================================= */

function doGet() {

  return jsonResponse({
    success: true,
    message: 'Shivam Gift House Billing API is running'
  });

}


/* =========================================================
   POST
   ========================================================= */

function doPost(e) {

  try {

    if (!e || !e.postData || !e.postData.contents) {

      return jsonResponse({
        success: false,
        error: 'No POST data received'
      });

    }


    const data =
      JSON.parse(e.postData.contents);


    /* SAVE BILL */

    if (data.action === 'saveBill') {

      return saveBill(data.bill);

    }


    /* SEARCH BILL HISTORY */

    if (data.action === 'searchBills') {

      /*
       * Support both formats:
       *
       * {
       *   action: "searchBills",
       *   search: {...}
       * }
       *
       * and
       *
       * {
       *   action: "searchBills",
       *   mode: "...",
       *   date: "..."
       * }
       */

      const search =
        data.search &&
        typeof data.search === 'object'
          ? data.search
          : data;


      return searchBills(search);

    }


    /* TODAY'S OVERVIEW */

    if (data.action === 'todayBills') {

      return getTodayBills();

    }


    return jsonResponse({

      success: false,

      error:
        'Unknown action: ' +
        (data.action || 'undefined')

    });


  } catch (error) {

    return jsonResponse({

      success: false,

      error:
        error.toString()

    });

  }

}


/* =========================================================
   SAVE BILL
   ========================================================= */

function saveBill(bill) {

  if (!bill) {

    return jsonResponse({

      success: false,

      error:
        'Bill data is missing'

    });

  }


  /* ITEMS ARE REQUIRED */

  if (
    !Array.isArray(bill.items) ||
    bill.items.length === 0
  ) {

    return jsonResponse({

      success: false,

      error:
        'Bill contains no items'

    });

  }


  /* PAYMENT IS REQUIRED */

  const paymentMode =
    bill.paymentMode ||
    bill.payment ||
    '';


  if (!paymentMode) {

    return jsonResponse({

      success: false,

      error:
        'Payment mode is required'

    });

  }


  const ss =
    SpreadsheetApp.openById(
      SHEET_ID
    );


  const date =
    new Date(
      bill.iso ||
      bill.date ||
      new Date()
    );


  if (isNaN(date.getTime())) {

    return jsonResponse({

      success: false,

      error:
        'Invalid bill date'

    });

  }


  const timezone =
    Session.getScriptTimeZone();


  const day =
    date.getDate();


  const monthName =
    Utilities.formatDate(
      date,
      timezone,
      'yyyy-MM'
    );


  let sheetName;


  if (day <= 15) {

    sheetName =
      monthName +
      '-01-15';

  } else {

    sheetName =
      monthName +
      '-16-END';

  }


  let sheet =
    ss.getSheetByName(
      sheetName
    );


  /* CREATE SHEET IF REQUIRED */

  if (!sheet) {

    sheet =
      ss.insertSheet(
        sheetName
      );


    sheet.appendRow([

      'Bill No.',
      'Date',
      'Time',

      'Customer Name',
      'Customer Phone',
      'Customer Address',

      'Particular',
      'Quantity',
      'Unit Price',
      'Item Total',

      'Subtotal',

      'Discount Type',
      'Discount',

      'Tax Label',
      'Tax Rate',
      'Tax Amount',

      'Grand Total',

      'Payment Mode',

      'Created By',

      'Currency'

    ]);


    sheet
      .getRange(
        1,
        1,
        1,
        20
      )
      .setFontWeight('bold');


    sheet.setFrozenRows(1);

  }


  /* BILL INFORMATION */

  const billNo =
    bill.billNo ||
    bill.id ||
    '';


  const customerName =
    bill.customerName ||
    bill.customer ||
    '';


  const customerPhone =
    bill.customerPhone ||
    bill.phone ||
    '';


  const customerAddress =
    bill.customerAddress ||
    bill.address ||
    '';


  const createdBy =
    bill.createdBy ||
    bill.user ||
    '';


  const currency =
    bill.currency ||
    'NPR';


  /* ADD EACH ITEM */

  bill.items.forEach(
    function(item) {

      const quantity =
        Number(
          item.quantity ??
          item.q ??
          0
        );


      const unitPrice =
        Number(
          item.unitPrice ??
          item.u ??
          0
        );


      const itemTotal =
        quantity *
        unitPrice;


      sheet.appendRow([

        billNo,

        Utilities.formatDate(
          date,
          timezone,
          'dd-MM-yyyy'
        ),

        Utilities.formatDate(
          date,
          timezone,
          'HH:mm:ss'
        ),

        customerName,

        customerPhone,

        customerAddress,

        item.particular ??
        item.p ??
        '',

        quantity,

        unitPrice,

        itemTotal,

        Number(
          bill.subtotal ||
          0
        ),

        bill.discountType ||
        'amount',

        Number(
          bill.discount ||
          0
        ),

        bill.taxLabel ||
        '',

        Number(
          bill.taxRate ||
          0
        ),

        Number(
          bill.taxAmount ||
          0
        ),

        Number(
          bill.grandTotal ||
          0
        ),

        paymentMode,

        createdBy,

        currency

      ]);

    }
  );


  sheet.autoResizeColumns(
    1,
    20
  );


  return jsonResponse({

    success: true,

    message:
      'Bill saved successfully',

    billNo:
      billNo

  });

}


/* =========================================================
   SEARCH BILL HISTORY
   ========================================================= */

function searchBills(search) {

  search =
    search ||
    {};


  const ss =
    SpreadsheetApp.openById(
      SHEET_ID
    );


  /*
   * Supported modes:
   *
   * date
   * range
   * today
   */

  const mode =
    String(
      search.mode ||
      'date'
    );


  const query =
    String(
      search.query ||
      search.customer ||
      ''
    )
      .trim()
      .toLowerCase();


  const billQuery =
    String(
      search.billNo ||
      ''
    )
      .trim()
      .toLowerCase();


  const payment =
    String(
      search.payment ||
      search.paymentMode ||
      'All'
    );


  let startDate;
  let endDate;


  /* TODAY */

  if (mode === 'today') {

    const now =
      new Date();


    startDate =
      new Date(
        now.getFullYear(),
        now.getMonth(),
        now.getDate()
      );


    endDate =
      new Date(
        now.getFullYear(),
        now.getMonth(),
        now.getDate()
      );

  }


  /* PARTICULAR DATE */

  else if (mode === 'date') {

    if (!search.date) {

      return jsonResponse({

        success: false,

        error:
          'Date is required'

      });

    }


    startDate =
      parseInputDate(
        search.date
      );


    endDate =
      parseInputDate(
        search.date
      );

  }


  /* DATE RANGE */

  else if (
    mode === 'range' ||
    mode === 'between'
  ) {

    if (
      !search.fromDate ||
      !search.toDate
    ) {

      return jsonResponse({

        success: false,

        error:
          'Both from and to dates are required'

      });

    }


    startDate =
      parseInputDate(
        search.fromDate
      );


    endDate =
      parseInputDate(
        search.toDate
      );

  }


  else {

    return jsonResponse({

      success: false,

      error:
        'Invalid search mode'

    });

  }


  if (
    !startDate ||
    !endDate ||
    isNaN(startDate.getTime()) ||
    isNaN(endDate.getTime())
  ) {

    return jsonResponse({

      success: false,

      error:
        'Invalid search date'

    });

  }


  startDate.setHours(
    0,
    0,
    0,
    0
  );


  endDate.setHours(
    23,
    59,
    59,
    999
  );


  /*
   * Only use our half-month sheets.
   */

  const sheets =
    ss
      .getSheets()
      .filter(
        function(sheet) {

          return /^\d{4}-\d{2}-(01-15|16-END)$/
            .test(
              sheet.getName()
            );

        }
      );


  const grouped = {};


  sheets.forEach(
    function(sheet) {

      const values =
        sheet
          .getDataRange()
          .getValues();


      if (values.length <= 1) {

        return;

      }


      for (
        let r = 1;
        r < values.length;
        r++
      ) {

        const row =
          values[r];


        const billNo =
          String(
            row[0] ||
            ''
          ).trim();


        if (!billNo) {

          continue;

        }


        const rowDate =
          parseSheetDate(
            row[1]
          );


        if (
          !rowDate ||
          rowDate < startDate ||
          rowDate > endDate
        ) {

          continue;

        }


        const customer =
          String(
            row[3] ||
            ''
          );


        const rowPayment =
          String(
            row[17] ||
            ''
          );


        /*
         * Search customer name.
         */

        if (
          query &&
          !customer
            .toLowerCase()
            .includes(query)
        ) {

          continue;

        }


        /*
         * Search bill number.
         */

        if (
          billQuery &&
          !billNo
            .toLowerCase()
            .includes(billQuery)
        ) {

          continue;

        }


        /*
         * Payment filter.
         */

        if (
          payment !== 'All' &&
          rowPayment.toLowerCase() !==
          payment.toLowerCase()
        ) {

          continue;

        }


        /*
         * Create bill group.
         */

        if (!grouped[billNo]) {

          grouped[billNo] = {

            id:
              billNo,

            date:
              Utilities.formatDate(
                rowDate,
                Session.getScriptTimeZone(),
                'dd-MM-yyyy'
              ),

            time:
              String(
                row[2] ||
                ''
              ),

            customer:
              customer,

            phone:
              String(
                row[4] ||
                ''
              ),

            address:
              String(
                row[5] ||
                ''
              ),

            payment:
              rowPayment,

            user:
              String(
                row[18] ||
                ''
              ),

            currency:
              String(
                row[19] ||
                'NPR'
              ),

            subtotal:
              Number(
                row[10] ||
                0
              ),

            discountType:
              String(
                row[11] ||
                'amount'
              ),

            discount:
              Number(
                row[12] ||
                0
              ),

            taxLabel:
              String(
                row[13] ||
                ''
              ),

            taxRate:
              Number(
                row[14] ||
                0
              ),

            taxAmount:
              Number(
                row[15] ||
                0
              ),

            grandTotal:
              Number(
                row[16] ||
                0
              ),

            items: []

          };

        }


        /*
         * Add item.
         */

        grouped[billNo]
          .items
          .push({

            p:
              String(
                row[6] ||
                ''
              ),

            q:
              Number(
                row[7] ||
                0
              ),

            u:
              Number(
                row[8] ||
                0
              )

          });

      }

    }
  );


  const bills =
    Object.keys(grouped)
      .map(
        function(key) {

          return grouped[key];

        }
      );


  /*
   * Sort newest first.
   */

  bills.sort(
    function(a, b) {

      const dateA =
        parseSheetDate(
          a.date
        );


      const dateB =
        parseSheetDate(
          b.date
        );


      if (
        dateB - dateA !== 0
      ) {

        return dateB - dateA;

      }


      return String(
        b.time
      ).localeCompare(
        String(a.time)
      );

    }
  );


  return jsonResponse({

    success: true,

    bills:
      bills,

    count:
      bills.length,

    message:
      bills.length === 0
        ? 'No data found'
        : 'Data found'

  });

}


/* =========================================================
   TODAY'S BILL OVERVIEW
   ========================================================= */

function getTodayBills() {

  const result =
    searchBills({

      mode:
        'today'

    });


  return result;

}


/* =========================================================
   DATE PARSER
   ========================================================= */

function parseInputDate(value) {

  if (!value) {

    return null;

  }


  const text =
    String(value)
      .trim();


  /*
   * HTML date input:
   * YYYY-MM-DD
   */

  const parts =
    text
      .split('-')
      .map(Number);


  if (
    parts.length !== 3 ||
    parts.some(
      function(n) {
        return isNaN(n);
      }
    )
  ) {

    return null;

  }


  return new Date(
    parts[0],
    parts[1] - 1,
    parts[2]
  );

}


/* =========================================================
   GOOGLE SHEET DATE PARSER
   ========================================================= */

function parseSheetDate(value) {

  /*
   * Google Sheets may return a Date object.
   */

  if (
    Object.prototype.toString.call(
      value
    ) === '[object Date]'
  ) {

    return new Date(
      value.getTime()
    );

  }


  const text =
    String(
      value ||
      ''
    ).trim();


  /*
   * Expected:
   *
   * DD-MM-YYYY
   */

  const m =
    text.match(
      /^(\d{2})-(\d{2})-(\d{4})$/
    );


  if (!m) {

    return null;

  }


  return new Date(

    Number(m[3]),

    Number(m[2]) - 1,

    Number(m[1])

  );

}


/* =========================================================
   JSON RESPONSE
   ========================================================= */

function jsonResponse(data) {

  return ContentService

    .createTextOutput(
      JSON.stringify(data)
    )

    .setMimeType(
      ContentService.MimeType.JSON
    );

}
