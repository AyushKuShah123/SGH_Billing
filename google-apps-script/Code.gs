const SHEET_ID = '1g-StRRD-sm-UWpyQhqTVodKEdmbUrC4vspKxKQoDiYs';

const SHEET_HEADER = [
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
];


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

    if (
      !e ||
      !e.postData ||
      !e.postData.contents
    ) {

      return jsonResponse({
        success: false,
        error: 'No POST data received'
      });

    }


    const data =
      JSON.parse(
        e.postData.contents
      );


    if (
      data.action === 'saveBill'
    ) {

      return saveBill(
        data.bill
      );

    }


    if (
      data.action === 'searchBills'
    ) {

      const search =
        data.search &&
        typeof data.search === 'object'
          ? data.search
          : data;

      return searchBills(
        search
      );

    }


    if (
      data.action === 'todayBills'
    ) {

      return getTodayBills();

    }


    return jsonResponse({

      success: false,

      error:
        'Unknown action: ' +
        (
          data.action ||
          'undefined'
        )

    });

  } catch (error) {

    return jsonResponse({

      success: false,

      error:
        error &&
        error.message
          ? error.message
          : String(error)

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


  if (
    !Array.isArray(
      bill.items
    ) ||
    bill.items.length === 0
  ) {

    return jsonResponse({

      success: false,

      error:
        'Bill contains no items'

    });

  }


  const paymentMode =
    String(
      bill.paymentMode ||
      bill.payment ||
      ''
    ).trim();


  if (!paymentMode) {

    return jsonResponse({

      success: false,

      error:
        'Payment mode is required'

    });

  }


  /*
   * BILL NUMBER IS THE UNIQUE IDENTIFIER.
   *
   * The same bill number can never be inserted twice.
   */

  const billNo =
    String(
      bill.billNo ||
      bill.id ||
      ''
    ).trim();


  if (!billNo) {

    return jsonResponse({

      success: false,

      error:
        'Bill number is missing'

    });

  }


  /*
   * Lock the script.
   *
   * This prevents two simultaneous requests from
   * checking for the bill at the same time and both
   * inserting it.
   */

  const lock =
    LockService.getScriptLock();


  try {

    lock.waitLock(30000);

  } catch (error) {

    return jsonResponse({

      success: false,

      error:
        'Server is busy. Please try again.'

    });

  }


  try {

    const ss =
      SpreadsheetApp.openById(
        SHEET_ID
      );


    /*
     * DUPLICATE CHECK
     *
     * This is performed BEFORE writing anything.
     */

    const existing =
      findBillByNumber(
        ss,
        billNo
      );


    if (existing) {

      return jsonResponse({

        success: true,

        duplicate: true,

        alreadySaved: true,

        message:
          'Bill already exists',

        billNo:
          billNo

      });

    }


    /*
     * BILL DATE
     */

    const date =
      getBillDate(
        bill
      );


    if (
      !date ||
      isNaN(
        date.getTime()
      )
    ) {

      return jsonResponse({

        success: false,

        error:
          'Invalid bill date'

      });

    }


    const timezone =
      Session.getScriptTimeZone();


    /*
     * Determine the correct billing sheet.
     */

    const day =
      Number(
        Utilities.formatDate(
          date,
          timezone,
          'dd'
        )
      );


    const monthName =
      Utilities.formatDate(
        date,
        timezone,
        'yyyy-MM'
      );


    const sheetName =
      day <= 15
        ? monthName + '-01-15'
        : monthName + '-16-END';


    let sheet =
      ss.getSheetByName(
        sheetName
      );


    /*
     * Create the sheet if necessary.
     */

    if (!sheet) {

      sheet =
        ss.insertSheet(
          sheetName
        );


      sheet
        .getRange(
          1,
          1,
          1,
          SHEET_HEADER.length
        )
        .setValues([
          SHEET_HEADER
        ]);


      sheet
        .getRange(
          1,
          1,
          1,
          SHEET_HEADER.length
        )
        .setFontWeight(
          'bold'
        );


      sheet.setFrozenRows(
        1
      );

    }


    /*
     * BILL INFORMATION
     */

    const customerName =
      String(
        bill.customerName ||
        bill.customer ||
        ''
      ).trim();


    const customerPhone =
      String(
        bill.customerPhone ||
        bill.phone ||
        ''
      ).trim();


    const customerAddress =
      String(
        bill.customerAddress ||
        bill.address ||
        ''
      ).trim();


    const createdBy =
      String(
        bill.createdBy ||
        bill.user ||
        ''
      ).trim();


    const currency =
      String(
        bill.currency ||
        'NPR'
      ).trim();


    const subtotal =
      toNumber(
        bill.subtotal
      );


    const discountType =
      String(
        bill.discountType ||
        'amount'
      ).trim();


    const discount =
      toNumber(
        bill.discount
      );


    const taxLabel =
      String(
        bill.taxLabel ||
        ''
      ).trim();


    const taxRate =
      toNumber(
        bill.taxRate
      );


    const taxAmount =
      toNumber(
        bill.taxAmount
      );


    const grandTotal =
      toNumber(
        bill.grandTotal
      );


    /*
     * Prepare all item rows first.
     */

    const rows = [];


    bill.items.forEach(
      function(item) {

        const quantity =
          toNumber(
            item.quantity ??
            item.q
          );


        const unitPrice =
          toNumber(
            item.unitPrice ??
            item.u
          );


        const itemTotal =
          quantity *
          unitPrice;


        const particular =
          String(
            item.particular ??
            item.p ??
            ''
          ).trim();


        rows.push([

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

          particular,

          quantity,

          unitPrice,

          itemTotal,

          subtotal,

          discountType,

          discount,

          taxLabel,

          taxRate,

          taxAmount,

          grandTotal,

          paymentMode,

          createdBy,

          currency

        ]);

      }
    );


    /*
     * Write all rows at once.
     */

    if (
      rows.length > 0
    ) {

      const firstRow =
        sheet.getLastRow() + 1;


      sheet
        .getRange(
          firstRow,
          1,
          rows.length,
          SHEET_HEADER.length
        )
        .setValues(
          rows
        );

    }


    /*
     * Make sure changes are committed before
     * returning the success response.
     */

    SpreadsheetApp.flush();


    return jsonResponse({

      success: true,

      duplicate: false,

      alreadySaved: false,

      message:
        'Bill saved successfully',

      billNo:
        billNo

    });


  } catch (error) {

    return jsonResponse({

      success: false,

      error:
        error &&
        error.message
          ? error.message
          : String(error)

    });


  } finally {

    try {

      lock.releaseLock();

    } catch (e) {}

  }

}


/* =========================================================
   FIND BILL BY BILL NUMBER
   ========================================================= */

function findBillByNumber(
  ss,
  billNo
) {

  const target =
    String(
      billNo
    )
    .trim()
    .toLowerCase();


  if (!target) {
    return null;
  }


  const sheets =
    getBillingSheets(
      ss
    );


  for (
    let i = 0;
    i < sheets.length;
    i++
  ) {

    const sheet =
      sheets[i];


    const lastRow =
      sheet.getLastRow();


    if (
      lastRow <= 1
    ) {

      continue;

    }


    /*
     * Bill number is column 1.
     */

    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getValues();


    for (
      let r = 0;
      r < values.length;
      r++
    ) {

      const existingBillNo =
        String(
          values[r][0] ||
          ''
        )
        .trim()
        .toLowerCase();


      if (
        existingBillNo ===
        target
      ) {

        return {

          sheet:
            sheet,

          row:
            r + 2

        };

      }

    }

  }


  return null;

}


/* =========================================================
   SEARCH BILL HISTORY
   ========================================================= */

function searchBills(search) {

  search =
    search || {};


  const ss =
    SpreadsheetApp.openById(
      SHEET_ID
    );


  const mode =
    String(
      search.mode ||
      'date'
    )
    .trim()
    .toLowerCase();


  const query =
    String(
      search.query ||
      ''
    )
    .trim()
    .toLowerCase();


  const billNoQuery =
    String(
      search.billNo ||
      ''
    )
    .trim()
    .toLowerCase();


  const payment =
    String(
      search.payment ||
      'All'
    )
    .trim();


  let startDate =
    null;

  let endDate =
    null;


  /*
   * TODAY
   */

  if (
    mode === 'today'
  ) {

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


  /*
   * SINGLE DATE
   */

  else if (
    mode === 'date'
  ) {

    if (
      !search.date
    ) {

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


  /*
   * DATE RANGE
   */

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


  /*
   * INVALID MODE
   */

  else {

    return jsonResponse({

      success: false,

      error:
        'Invalid search mode'

    });

  }


  /*
   * Validate dates.
   */

  if (
    !startDate ||
    !endDate ||
    isNaN(
      startDate.getTime()
    ) ||
    isNaN(
      endDate.getTime()
    )
  ) {

    return jsonResponse({

      success: false,

      error:
        'Invalid search date'

    });

  }


  /*
   * Compare dates using date-only values.
   */

  startDate =
    dateOnly(
      startDate
    );


  endDate =
    dateOnly(
      endDate
    );


  /*
   * Get only the billing sheets.
   */

  const sheets =
    getBillingSheets(
      ss
    );


  const grouped = {};


  sheets.forEach(
    function(sheet) {

      const values =
        sheet
          .getDataRange()
          .getValues();


      if (
        !values ||
        values.length <= 1
      ) {

        return;

      }


      for (
        let r = 1;
        r < values.length;
        r++
      ) {

        const row =
          values[r];


        /*
         * Column 0 = Bill No.
         */

        const billNo =
          String(
            row[0] ||
            ''
          ).trim();


        if (!billNo) {

          continue;

        }


        /*
         * Column 1 = Date
         */

        const rowDate =
          parseSheetDate(
            row[1]
          );


        if (!rowDate) {

          continue;

        }


        const comparableDate =
          dateOnly(
            rowDate
          );


        /*
         * Date filter.
         */

        if (
          comparableDate < startDate ||
          comparableDate > endDate
        ) {

          continue;

        }


        /*
         * Column 3 = Customer Name
         */

        const customer =
          String(
            row[3] ||
            ''
          ).trim();


        /*
         * Column 17 = Payment Mode
         */

        const rowPayment =
          String(
            row[17] ||
            ''
          ).trim();


        /*
         * Customer search.
         */

        if (
          query &&
          !customer
            .toLowerCase()
            .includes(
              query
            )
        ) {

          continue;

        }


        /*
         * Bill number search.
         */

        if (
          billNoQuery &&
          !billNo
            .toLowerCase()
            .includes(
              billNoQuery
            )
        ) {

          continue;

        }


        /*
         * Payment filter.
         */

        if (
          payment &&
          payment.toLowerCase() !==
          'all' &&
          rowPayment.toLowerCase() !==
          payment.toLowerCase()
        ) {

          continue;

        }


        /*
         * GROUP BILL ITEMS
         *
         * Every item of the same bill number
         * is combined into one bill object.
         */

        if (
          !grouped[billNo]
        ) {

          grouped[billNo] = {

            id:
              billNo,

            billNo:
              billNo,

            date:
              formatDisplayDate(
                rowDate
              ),

            time:
              formatDisplayTime(
                row[2]
              ),

            customer:
              customer,

            customerName:
              customer,

            phone:
              String(
                row[4] ||
                ''
              ).trim(),

            customerPhone:
              String(
                row[4] ||
                ''
              ).trim(),

            address:
              String(
                row[5] ||
                ''
              ).trim(),

            customerAddress:
              String(
                row[5] ||
                ''
              ).trim(),

            payment:
              rowPayment,

            paymentMode:
              rowPayment,

            user:
              String(
                row[18] ||
                ''
              ).trim(),

            createdBy:
              String(
                row[18] ||
                ''
              ).trim(),

            currency:
              String(
                row[19] ||
                'NPR'
              ).trim(),

            subtotal:
              toNumber(
                row[10]
              ),

            discountType:
              String(
                row[11] ||
                'amount'
              ).trim(),

            discount:
              toNumber(
                row[12]
              ),

            taxLabel:
              String(
                row[13] ||
                ''
              ).trim(),

            taxRate:
              toNumber(
                row[14]
              ),

            taxAmount:
              toNumber(
                row[15]
              ),

            grandTotal:
              toNumber(
                row[16]
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
              ).trim(),

            particular:
              String(
                row[6] ||
                ''
              ).trim(),

            q:
              toNumber(
                row[7]
              ),

            quantity:
              toNumber(
                row[7]
              ),

            u:
              toNumber(
                row[8]
              ),

            unitPrice:
              toNumber(
                row[8]
              ),

            itemTotal:
              toNumber(
                row[9]
              )

          });

      }

    }
  );


  /*
   * Convert grouped object to array.
   */

  const bills =
    Object.keys(
      grouped
    ).map(
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
        dateB &&
        dateA &&
        dateB.getTime() !==
        dateA.getTime()
      ) {

        return (
          dateB.getTime() -
          dateA.getTime()
        );

      }


      return String(
        b.time ||
        ''
      ).localeCompare(
        String(
          a.time ||
          ''
        )
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
      bills.length > 0
        ? 'Data found'
        : 'No data found'

  });

}


/* =========================================================
   TODAY'S BILL OVERVIEW
   ========================================================= */

function getTodayBills() {

  return searchBills({

    mode:
      'today'

  });

}


/* =========================================================
   GET BILLING SHEETS
   ========================================================= */

function getBillingSheets(ss) {

  return ss
    .getSheets()
    .filter(
      function(sheet) {

        return /^\d{4}-\d{2}-(01-15|16-END)$/
          .test(
            sheet.getName()
          );

      }
    );

}


/* =========================================================
   INPUT DATE PARSER
   ========================================================= */

function parseInputDate(value) {

  if (!value) {

    return null;

  }


  const text =
    String(
      value
    ).trim();


  /*
   * Expected:
   *
   * YYYY-MM-DD
   */

  const match =
    text.match(
      /^(\d{4})-(\d{2})-(\d{2})$/
    );


  if (!match) {

    return null;

  }


  const year =
    Number(
      match[1]
    );


  const month =
    Number(
      match[2]
    );


  const day =
    Number(
      match[3]
    );


  const date =
    new Date(
      year,
      month - 1,
      day
    );


  /*
   * Reject impossible dates.
   */

  if (
    date.getFullYear() !==
      year ||
    date.getMonth() !==
      month - 1 ||
    date.getDate() !==
      day
  ) {

    return null;

  }


  return date;

}


/* =========================================================
   SHEET DATE PARSER
   ========================================================= */

function parseSheetDate(value) {

  /*
   * Native Date object.
   */

  if (
    Object.prototype.toString.call(
      value
    ) === '[object Date]'
  ) {

    if (
      isNaN(
        value.getTime()
      )
    ) {

      return null;

    }


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
   * DD-MM-YYYY
   */

  let match =
    text.match(
      /^(\d{1,2})-(\d{1,2})-(\d{4})$/
    );


  if (match) {

    const day =
      Number(
        match[1]
      );

    const month =
      Number(
        match[2]
      );

    const year =
      Number(
        match[3]
      );


    const date =
      new Date(
        year,
        month - 1,
        day
      );


    if (
      date.getFullYear() !==
        year ||
      date.getMonth() !==
        month - 1 ||
      date.getDate() !==
        day
    ) {

      return null;

    }


    return date;

  }


  /*
   * YYYY-MM-DD
   */

  match =
    text.match(
      /^(\d{4})-(\d{1,2})-(\d{1,2})$/
    );


  if (match) {

    const year =
      Number(
        match[1]
      );

    const month =
      Number(
        match[2]
      );

    const day =
      Number(
        match[3]
      );


    const date =
      new Date(
        year,
        month - 1,
        day
      );


    if (
      date.getFullYear() !==
        year ||
      date.getMonth() !==
        month - 1 ||
      date.getDate() !==
        day
    ) {

      return null;

    }


    return date;

  }


  return null;

}


/* =========================================================
   GET BILL DATE
   ========================================================= */

function getBillDate(bill) {

  const candidates = [

    bill.date,

    bill.iso,

    bill.createdAt,

    bill.timestamp

  ];


  for (
    let i = 0;
    i < candidates.length;
    i++
  ) {

    const value =
      candidates[i];


    if (
      value === undefined ||
      value === null ||
      value === ''
    ) {

      continue;

    }


    /*
     * If already a Date object.
     */

    if (
      Object.prototype.toString.call(
        value
      ) === '[object Date]'
    ) {

      if (
        !isNaN(
          value.getTime()
        )
      ) {

        return new Date(
          value.getTime()
        );

      }

      continue;

    }


    const text =
      String(
        value
      ).trim();


    /*
     * YYYY-MM-DD
     */

    if (
      /^\d{4}-\d{2}-\d{2}$/.test(
        text
      )
    ) {

      const parsed =
        parseInputDate(
          text
        );


      if (parsed) {

        return parsed;

      }

    }


    /*
     * Try normal JavaScript date parsing.
     */

    const parsed =
      new Date(
        text
      );


    if (
      !isNaN(
        parsed.getTime()
      )
    ) {

      return parsed;

    }

  }


  /*
   * If no valid date was supplied,
   * use current server date.
   */

  return new Date();

}


/* =========================================================
   DATE ONLY
   ========================================================= */

function dateOnly(date) {

  return new Date(

    date.getFullYear(),

    date.getMonth(),

    date.getDate(),

    0,
    0,
    0,
    0

  );

}


/* =========================================================
   DISPLAY DATE
   ========================================================= */

function formatDisplayDate(date) {

  if (
    !date ||
    isNaN(
      date.getTime()
    )
  ) {

    return '';

  }


  const timezone =
    Session.getScriptTimeZone();


  return Utilities.formatDate(

    date,

    timezone,

    'dd-MM-yyyy'

  );

}


/* =========================================================
   DISPLAY TIME
   ========================================================= */

function formatDisplayTime(value) {

  if (
    Object.prototype.toString.call(
      value
    ) === '[object Date]'
  ) {

    if (
      isNaN(
        value.getTime()
      )
    ) {

      return '';

    }


    return Utilities.formatDate(

      value,

      Session.getScriptTimeZone(),

      'HH:mm:ss'

    );

  }


  return String(
    value ||
    ''
  ).trim();

}


/* =========================================================
   NUMBER CONVERTER
   ========================================================= */

function toNumber(value) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {

    return 0;

  }


  const number =
    Number(
      value
    );


  return isNaN(
    number
  )
    ? 0
    : number;

}


/* =========================================================
   JSON RESPONSE
   ========================================================= */

function jsonResponse(data) {

  return ContentService

    .createTextOutput(
      JSON.stringify(
        data
      )
    )

    .setMimeType(
      ContentService.MimeType.JSON
    );

}
