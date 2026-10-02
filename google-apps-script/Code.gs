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

const BILL_PREFIX_PROPERTY = 'SGH_BILL_SEQ_';
const BILL_SEPARATOR_ROWS = 1;

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

    const data = JSON.parse(e.postData.contents);

    if (data.action === 'saveBill') {
      return saveBill(data.bill);
    }

    if (data.action === 'searchBills') {
      const search =
        data.search && typeof data.search === 'object'
          ? data.search
          : data;
      return searchBills(search);
    }

    if (data.action === 'todayBills') {
      return getTodayBills();
    }

    if (data.action === 'nextBillNo') {
      return getNextBillNo();
    }

    return jsonResponse({
      success: false,
      error: 'Unknown action: ' + (data.action || 'undefined')
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      error: error && error.message ? error.message : String(error)
    });
  }
}

/* =========================================================
   SERVER-SIDE BILL NUMBER
   ---------------------------------------------------------
   Bill number is generated from the spreadsheet, not from
   browser memory. This means refreshing the browser does
   NOT reset the sequence.

   Format:
   SGH-YYDAYMONTHLETTER0001

   Example:
   SGH-2601A0001

   The sequence is maintained per prefix in Script Properties.
   A LockService lock prevents two users from receiving the
   same number at the same time.
   ========================================================= */

function getBillPrefix_(date) {
  const timezone = Session.getScriptTimeZone();

  const yy = Utilities.formatDate(date, timezone, 'yy');
  const day = Utilities.formatDate(date, timezone, 'dd');
  const monthNumber = Number(Utilities.formatDate(date, timezone, 'M'));

  const monthLetter = String.fromCharCode(64 + monthNumber);

  return 'SGH-' + yy + day + monthLetter;
}

function getNextBillNo() {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const now = new Date();
    const prefix = getBillPrefix_(now);

    const nextNumber = getNextSequenceNumber_(ss, prefix);
    const billNo = prefix + String(nextNumber).padStart(4, '0');

    return jsonResponse({
      success: true,
      billNo: billNo,
      prefix: prefix,
      sequence: nextNumber,
      message: 'Next bill number calculated'
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      error: error && error.message ? error.message : String(error)
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

/*
 * Find the largest existing number in the spreadsheet and
 * compare it with the saved script-property sequence.
 *
 * This self-heals if rows were manually deleted/added.
 */
function getNextSequenceNumber_(ss, prefix) {
  const props = PropertiesService.getScriptProperties();
  const propertyKey = BILL_PREFIX_PROPERTY + prefix;

  let propertyValue = Number(props.getProperty(propertyKey) || 0);
  let sheetMax = 0;

  const sheets = getBillingSheets(ss);

  sheets.forEach(function(sheet) {
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) return;

    const values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();

    values.forEach(function(row) {
      const value = String(row[0] || '').trim();

      if (!value || value.indexOf(prefix) !== 0) return;

      const suffix = value.slice(prefix.length);
      const n = Number(suffix);

      if (Number.isFinite(n) && n > sheetMax) {
        sheetMax = n;
      }
    });
  });

  const next = Math.max(propertyValue, sheetMax) + 1;

  props.setProperty(propertyKey, String(next));

  return next;
}

/* =========================================================
   SAVE BILL
   ========================================================= */

function saveBill(bill) {
  if (!bill) {
    return jsonResponse({
      success: false,
      error: 'Bill data is missing'
    });
  }

  if (!Array.isArray(bill.items) || bill.items.length === 0) {
    return jsonResponse({
      success: false,
      error: 'Bill contains no items'
    });
  }

  const paymentMode = String(
    bill.paymentMode || bill.payment || ''
  ).trim();

  if (!paymentMode) {
    return jsonResponse({
      success: false,
      error: 'Payment mode is required'
    });
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);
  } catch (error) {
    return jsonResponse({
      success: false,
      error: 'Server is busy. Please try again.'
    });
  }

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    /*
     * Date is determined first so the server can generate the
     * correct prefix.
     */
    const date = new Date(
      bill.date ||
      bill.iso ||
      new Date()
    );

    if (isNaN(date.getTime())) {
      return jsonResponse({
        success: false,
        error: 'Invalid bill date'
      });
    }

    /*
     * SERVER-SIDE BILL NUMBER
     *
     * The browser no longer decides the final bill number.
     * If the browser sends an existing number, we first check
     * it for duplicate protection. Otherwise, generate one
     * atomically from the server.
     */
    let requestedBillNo = String(
      bill.billNo ||
      bill.id ||
      ''
    ).trim();

    if (requestedBillNo) {
      const existingRequested = findBillByNumber(ss, requestedBillNo);

      if (existingRequested) {
        return jsonResponse({
          success: true,
          duplicate: true,
          message: 'Bill already exists',
          billNo: requestedBillNo
        });
      }
    }

    const prefix = getBillPrefix_(date);

    /*
     * Always generate the final bill number on the server.
     * This guarantees continuity after browser refresh.
     */
    const nextNumber = getNextSequenceNumber_(ss, prefix);
    const billNo =
      prefix +
      String(nextNumber).padStart(4, '0');

    /*
     * Extra duplicate check after generation.
     * Normally impossible because of LockService, but this
     * keeps the function defensive.
     */
    const existing = findBillByNumber(ss, billNo);

    if (existing) {
      return jsonResponse({
        success: true,
        duplicate: true,
        message: 'Bill already exists',
        billNo: billNo
      });
    }

    const timezone = Session.getScriptTimeZone();

    const day = Number(
      Utilities.formatDate(date, timezone, 'dd')
    );

    const monthName = Utilities.formatDate(
      date,
      timezone,
      'yyyy-MM'
    );

    const sheetName =
      day <= 15
        ? monthName + '-01-15'
        : monthName + '-16-END';

    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);

      sheet
        .getRange(
          1,
          1,
          1,
          SHEET_HEADER.length
        )
        .setValues([SHEET_HEADER]);

      sheet
        .getRange(
          1,
          1,
          1,
          SHEET_HEADER.length
        )
        .setFontWeight('bold');

      sheet.setFrozenRows(1);
    } else {
      /*
       * Repair header if the sheet exists but is empty.
       */
      if (sheet.getLastRow() === 0) {
        sheet
          .getRange(
            1,
            1,
            1,
            SHEET_HEADER.length
          )
          .setValues([SHEET_HEADER]);

        sheet.setFrozenRows(1);
      }
    }

    const customerName = String(
      bill.customerName ||
      bill.customer ||
      ''
    );

    const customerPhone = String(
      bill.customerPhone ||
      bill.phone ||
      ''
    );

    const customerAddress = String(
      bill.customerAddress ||
      bill.address ||
      ''
    );

    const createdBy = String(
      bill.createdBy ||
      bill.user ||
      ''
    );

    const currency = String(
      bill.currency ||
      'NPR'
    );

    const subtotal = Number(
      bill.subtotal || 0
    );

    const discountType = String(
      bill.discountType ||
      'amount'
    );

    const discount = Number(
      bill.discount || 0
    );

    const taxLabel = String(
      bill.taxLabel ||
      ''
    );

    const taxRate = Number(
      bill.taxRate || 0
    );

    const taxAmount = Number(
      bill.taxAmount || 0
    );

    const grandTotal = Number(
      bill.grandTotal || 0
    );

    const rows = [];

    bill.items.forEach(function(item) {
      const quantity = Number(
        item.quantity ??
        item.q ??
        0
      );

      const unitPrice = Number(
        item.unitPrice ??
        item.u ??
        0
      );

      const itemTotal =
        quantity * unitPrice;

      const particular = String(
        item.particular ??
        item.p ??
        ''
      );

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
    });

    /*
     * Remove accidental trailing blank rows before adding the
     * new bill. Then write the complete bill.
     */
    const firstRow = Math.max(
      sheet.getLastRow() + 1,
      2
    );

    sheet
      .getRange(
        firstRow,
        1,
        rows.length,
        SHEET_HEADER.length
      )
      .setValues(rows);

    /*
     * One completely blank row between bills.
     *
     * Because every bill may contain multiple item rows,
     * the blank row is placed AFTER the last item of this bill.
     */
    const separatorRow = firstRow + rows.length;

    sheet
      .getRange(
        separatorRow,
        1,
        1,
        SHEET_HEADER.length
      )
      .clearContent();

    return jsonResponse({
      success: true,
      duplicate: false,
      message: 'Bill saved successfully',
      billNo: billNo,
      sequence: nextNumber,
      sheet: sheetName
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      error: error && error.message
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
   FIND EXISTING BILL
   ========================================================= */

function findBillByNumber(ss, billNo) {
  const sheets = getBillingSheets(ss);

  const target = String(
    billNo
  )
    .trim()
    .toLowerCase();

  if (!target) return null;

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];

    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      continue;
    }

    const values = sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getDisplayValues();

    for (let r = 0; r < values.length; r++) {
      const value = String(
        values[r][0] || ''
      )
        .trim()
        .toLowerCase();

      if (value === target) {
        return {
          sheet: sheet,
          row: r + 2
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
  search = search || {};

  const ss = SpreadsheetApp.openById(SHEET_ID);

  const mode = String(
    search.mode || 'date'
  )
    .toLowerCase()
    .trim();

  const query = String(
    search.query || ''
  )
    .trim()
    .toLowerCase();

  const billNoQuery = String(
    search.billNo || ''
  )
    .trim()
    .toLowerCase();

  const payment = String(
    search.payment || 'All'
  ).trim();

  let startDate;
  let endDate;

  if (mode === 'today') {
    const today = new Date();

    startDate = new Date(
      today.getFullYear(),
      today.getMonth(),
      today.getDate()
    );

    endDate = new Date(
      today.getFullYear(),
      today.getMonth(),
      today.getDate()
    );

  } else if (mode === 'date') {
    if (!search.date) {
      return jsonResponse({
        success: false,
        error: 'Date is required'
      });
    }

    startDate = parseInputDate(search.date);
    endDate = parseInputDate(search.date);

  } else if (
    mode === 'range' ||
    mode === 'between'
  ) {
    if (
      !search.fromDate ||
      !search.toDate
    ) {
      return jsonResponse({
        success: false,
        error: 'Both from and to dates are required'
      });
    }

    startDate = parseInputDate(search.fromDate);
    endDate = parseInputDate(search.toDate);

  } else {
    return jsonResponse({
      success: false,
      error: 'Invalid search mode'
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
      error: 'Invalid search date'
    });
  }

  startDate.setHours(
    0, 0, 0, 0
  );

  endDate.setHours(
    23, 59, 59, 999
  );

  const sheets = getBillingSheets(ss);

  const grouped = {};

  sheets.forEach(function(sheet) {
    const values = sheet
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
      const row = values[r];

      /*
       * Blank separator row.
       */
      const billNo = String(
        row[0] || ''
      ).trim();

      if (!billNo) {
        continue;
      }

      const rowDate = parseSheetDate(row[1]);

      if (!rowDate) {
        continue;
      }

      const comparableDate = new Date(
        rowDate.getFullYear(),
        rowDate.getMonth(),
        rowDate.getDate()
      );

      if (
        comparableDate < startDate ||
        comparableDate > endDate
      ) {
        continue;
      }

      const customer = String(
        row[3] || ''
      ).trim();

      const rowPayment = String(
        row[17] || ''
      ).trim();

      /*
       * Customer search.
       *
       * Also allow bill number search through the same
       * customer/bill field so the UI can use one box.
       */
      if (
        query &&
        !customer.toLowerCase().includes(query) &&
        !billNo.toLowerCase().includes(query)
      ) {
        continue;
      }

      if (
        billNoQuery &&
        !billNo.toLowerCase().includes(billNoQuery)
      ) {
        continue;
      }

      if (
        payment !== 'All' &&
        rowPayment.toLowerCase() !==
        payment.toLowerCase()
      ) {
        continue;
      }

      if (!grouped[billNo]) {
        grouped[billNo] = {
          id: billNo,
          billNo: billNo,

          date: Utilities.formatDate(
            rowDate,
            Session.getScriptTimeZone(),
            'dd-MM-yyyy'
          ),

          time: String(row[2] || ''),

          customer: customer,
          customerName: customer,

          phone: String(row[4] || ''),
          address: String(row[5] || ''),

          payment: rowPayment,
          paymentMode: rowPayment,

          user: String(row[18] || ''),
          createdBy: String(row[18] || ''),

          currency: String(
            row[19] || 'NPR'
          ),

          subtotal: Number(row[10] || 0),

          discountType: String(
            row[11] || 'amount'
          ),

          discount: Number(row[12] || 0),

          taxLabel: String(row[13] || ''),
          taxRate: Number(row[14] || 0),
          taxAmount: Number(row[15] || 0),

          grandTotal: Number(row[16] || 0),

          items: []
        };
      }

      grouped[billNo].items.push({
        p: String(row[6] || ''),
        q: Number(row[7] || 0),
        u: Number(row[8] || 0)
      });
    }
  });

  const bills = Object.keys(grouped)
    .map(function(key) {
      return grouped[key];
    });

  /*
   * Sort by date and then time, newest first.
   */
  bills.sort(function(a, b) {
    const dateA = parseSheetDate(a.date);
    const dateB = parseSheetDate(b.date);

    const dateDifference = dateB - dateA;

    if (dateDifference !== 0) {
      return dateDifference;
    }

    return String(
      b.time || ''
    ).localeCompare(
      String(a.time || '')
    );
  });

  return jsonResponse({
    success: true,
    bills: bills,
    count: bills.length,
    message: bills.length
      ? 'Data found'
      : 'No data found'
  });
}

/* =========================================================
   TODAY
   ========================================================= */

function getTodayBills() {
  return searchBills({
    mode: 'today'
  });
}

/* =========================================================
   BILLING SHEETS
   ========================================================= */

function getBillingSheets(ss) {
  return ss
    .getSheets()
    .filter(function(sheet) {
      return /^\d{4}-\d{2}-(01-15|16-END)$/.test(
        sheet.getName()
      );
    });
}

/* =========================================================
   INPUT DATE PARSER
   ========================================================= */

function parseInputDate(value) {
  if (!value) return null;

  const text = String(value).trim();

  const match = text.match(
    /^(\d{4})-(\d{2})-(\d{2})$/
  );

  if (!match) return null;

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);

  const date = new Date(
    year,
    month - 1,
    day
  );

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }

  return date;
}

/* =========================================================
   SHEET DATE PARSER
   ========================================================= */

function parseSheetDate(value) {
  if (
    Object.prototype.toString.call(value) ===
    '[object Date]'
  ) {
    if (isNaN(value.getTime())) {
      return null;
    }

    return new Date(value.getTime());
  }

  const text = String(
    value || ''
  ).trim();

  let match = text.match(
    /^(\d{1,2})-(\d{1,2})-(\d{4})$/
  );

  if (match) {
    return new Date(
      Number(match[3]),
      Number(match[2]) - 1,
      Number(match[1])
    );
  }

  match = text.match(
    /^(\d{4})-(\d{1,2})-(\d{1,2})$/
  );

  if (match) {
    return new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3])
    );
  }

  return null;
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
