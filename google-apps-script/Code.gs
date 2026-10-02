const SHEET_ID = '1g-StRRD-sm-UWpyQhqTVodKEdmbUrC4vspKxKQoDiYs';

function doGet() {
  return jsonResponse({
    success: true,
    message: 'Shivam Gift House Billing API is running'
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ success: false, error: 'No POST data received' });
    }

    const data = JSON.parse(e.postData.contents);

    if (data.action === 'saveBill') {
      return saveBill(data.bill);
    }

    if (data.action === 'searchBills') {
      return searchBills(data);
    }

    return jsonResponse({
      success: false,
      error: 'Unknown action'
    });

  } catch (error) {
    return jsonResponse({
      success: false,
      error: error.toString()
    });
  }
}


function getBillingSheets(ss) {
  return ss.getSheets().filter(function(sheet) {
    return /^\d{4}-\d{2}-(01-15|16-END)$/.test(sheet.getName());
  });
}

function findBillByNumber(ss, billNo) {
  const target = String(billNo || '').trim().toLowerCase();
  if (!target) return null;

  const sheets = getBillingSheets(ss);

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;

    const values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();

    for (let r = 0; r < values.length; r++) {
      if (String(values[r][0] || '').trim().toLowerCase() === target) {
        return { sheet: sheet, row: r + 2 };
      }
    }
  }

  return null;
}

/*
 * Server-side sequential bill number.
 * Format: SGH-YYDDM0001
 *
 * The browser no longer decides the next number.
 * Script Properties + Script Lock keep the sequence stable
 * even when the browser is refreshed or multiple saves occur.
 */
function generateNextBillNumber(ss, date) {
  const timezone = Session.getScriptTimeZone();

  const yy = Utilities.formatDate(date, timezone, 'yy');
  const dd = Utilities.formatDate(date, timezone, 'dd');
  const monthIndex = Number(Utilities.formatDate(date, timezone, 'M')) - 1;
  const monthLetter = String.fromCharCode(65 + monthIndex);

  const prefix = 'SGH-' + yy + dd + monthLetter;
  const propertyKey = 'BILL_COUNTER_' + prefix;

  const props = PropertiesService.getScriptProperties();
  let current = Number(props.getProperty(propertyKey) || 0);

  // First use: recover the highest existing number from the sheet.
  if (current <= 0) {
    const escapedPrefix = prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp('^' + escapedPrefix + '(\\d+)$');

    getBillingSheets(ss).forEach(function(sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return;

      const values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();

      values.forEach(function(row) {
        const value = String(row[0] || '').trim();
        const match = value.match(regex);

        if (match) {
          current = Math.max(current, Number(match[1]) || 0);
        }
      });
    });
  }

  const next = current + 1;
  props.setProperty(propertyKey, String(next));

  return prefix + String(next).padStart(4, '0');
}


function saveBill(bill) {
  if (!bill) {
    return jsonResponse({ success: false, error: 'Bill data is missing' });
  }

  if (!Array.isArray(bill.items) || bill.items.length === 0) {
    return jsonResponse({ success: false, error: 'Bill contains no items' });
  }

  if (!bill.paymentMode && !bill.payment) {
    return jsonResponse({ success: false, error: 'Payment mode is required' });
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
    const date = new Date(bill.date || new Date());

    if (isNaN(date.getTime())) {
      return jsonResponse({ success: false, error: 'Invalid bill date' });
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

      sheet.getRange(1, 1, 1, 20).setValues([[
        'Bill No.', 'Date', 'Time', 'Customer Name', 'Customer Phone',
        'Customer Address', 'Particular', 'Quantity', 'Unit Price', 'Item Total',
        'Subtotal', 'Discount Type', 'Discount', 'Tax Label', 'Tax Rate',
        'Tax Amount', 'Grand Total', 'Payment Mode', 'Created By', 'Currency'
      ]]);

      sheet.getRange(1, 1, 1, 20).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // SERVER is authoritative for the bill number.
    const billNo = generateNextBillNumber(ss, date);

    // Final collision protection.
    if (findBillByNumber(ss, billNo)) {
      return jsonResponse({
        success: false,
        error: 'Generated bill number already exists. Please try again.'
      });
    }

    const customerName = String(
      bill.customerName || bill.customer || ''
    );

    const customerPhone = String(
      bill.customerPhone || bill.phone || ''
    );

    const customerAddress = String(
      bill.customerAddress || bill.address || ''
    );

    const paymentMode = String(
      bill.paymentMode || bill.payment || ''
    );

    const rows = bill.items.map(function(item) {
      const quantity = Number(
        item.quantity ?? item.q ?? 0
      );

      const unitPrice = Number(
        item.unitPrice ?? item.u ?? 0
      );

      return [
        billNo,
        Utilities.formatDate(date, timezone, 'dd-MM-yyyy'),
        Utilities.formatDate(date, timezone, 'HH:mm:ss'),
        customerName,
        customerPhone,
        customerAddress,
        item.particular ?? item.p ?? '',
        quantity,
        unitPrice,
        quantity * unitPrice,
        Number(bill.subtotal || 0),
        bill.discountType || 'amount',
        Number(bill.discount || 0),
        bill.taxLabel || '',
        Number(bill.taxRate || 0),
        Number(bill.taxAmount || 0),
        Number(bill.grandTotal || 0),
        paymentMode,
        bill.createdBy || bill.user || '',
        bill.currency || 'NPR'
      ];
    });

    /*
     * Save the whole bill in one operation.
     *
     * The next row is intentionally left EMPTY.
     * Therefore there is one blank row between bills.
     */
    const firstRow = Math.max(
      2,
      sheet.getLastRow() + 1
    );

    sheet.getRange(
      firstRow,
      1,
      rows.length,
      20
    ).setValues(rows);

    // Do not write anything to firstRow + rows.length.
    // That is the blank separator row.

    return jsonResponse({
      success: true,
      duplicate: false,
      message: 'Bill saved successfully',
      billNo: billNo
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

function searchBills(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const mode = data.mode || 'date';
  const query = String(data.query || '').trim().toLowerCase();
  const payment = String(data.payment || 'All');

  let startDate = null;
  let endDate = null;

  if (mode === 'date') {
    if (!data.date) return jsonResponse({ success: false, error: 'Date is required' });
    startDate = parseInputDate(data.date);
    endDate = new Date(startDate);
  } else {
    if (!data.fromDate || !data.toDate) {
      return jsonResponse({ success: false, error: 'Both from and to dates are required' });
    }
    startDate = parseInputDate(data.fromDate);
    endDate = parseInputDate(data.toDate);
  }

  if (!startDate || !endDate || isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
    return jsonResponse({ success: false, error: 'Invalid search date' });
  }

  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(23, 59, 59, 999);

  const sheets = getBillingSheets(ss);

  const grouped = {};

  sheets.forEach(function(sheet) {
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return;

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const billNo = String(row[0] || '').trim();
      if (!billNo) continue;

      const rowDate = parseSheetDate(row[1]);
      if (!rowDate || rowDate < startDate || rowDate > endDate) continue;

      const customer = String(row[3] || '');
      const rowPayment = String(row[17] || '');
      const haystack = (billNo + ' ' + customer).toLowerCase();

      if (query && !haystack.includes(query)) continue;
      if (payment !== 'All' && rowPayment !== payment) continue;

      if (!grouped[billNo]) {
        grouped[billNo] = {
          id: billNo,
          date: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'dd-MM-yyyy'),
          customer: customer,
          phone: String(row[4] || ''),
          address: String(row[5] || ''),
          payment: rowPayment,
          user: String(row[18] || ''),
          currency: String(row[19] || 'NPR'),
          subtotal: Number(row[10] || 0),
          discountType: String(row[11] || 'amount'),
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

  const bills = Object.keys(grouped).map(function(k) { return grouped[k]; });
  bills.sort(function(a, b) {
    return parseSheetDate(b.date) - parseSheetDate(a.date);
  });

  return jsonResponse({
    success: true,
    bills: bills,
    count: bills.length
  });
}

function parseInputDate(value) {
  const parts = String(value).split('-').map(Number);
  if (parts.length !== 3) return null;
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function parseSheetDate(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return new Date(value.getTime());
  }

  const text = String(value || '').trim();
  const m = text.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
