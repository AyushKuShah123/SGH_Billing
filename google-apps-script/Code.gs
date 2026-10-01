const SHEET_ID = '1g-StRRD-sm-UWpyQhqTVodKEdmbUrC4vspKxKQoDiYs';

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      message: 'Shivam Gift House Billing API is running'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'saveBill') {
      return saveBill(data.bill);
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


function saveBill(bill) {

  const ss = SpreadsheetApp.openById(SHEET_ID);

  const date = new Date(bill.date || new Date());

  const day = date.getDate();

  let sheetName;

  if (day <= 15) {
    sheetName = Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      'yyyy-MM'
    ) + '-01-15';
  } else {
    sheetName = Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      'yyyy-MM'
    ) + '-16-END';
  }

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);

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

    sheet.getRange(1, 1, 1, 20)
      .setFontWeight('bold');

    sheet.setFrozenRows(1);
  }

  const items = bill.items || [];

  items.forEach(function(item) {

    const itemTotal =
      Number(item.quantity || item.q || 0) *
      Number(item.unitPrice || item.u || 0);

    sheet.appendRow([
      bill.billNo || bill.id || '',
      Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd-MM-yyyy'),
      Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm:ss'),

      bill.customerName || bill.customer || '',
      bill.customerPhone || bill.phone || '',
      bill.customerAddress || bill.address || '',

      item.particular || item.p || '',
      Number(item.quantity || item.q || 0),
      Number(item.unitPrice || item.u || 0),
      itemTotal,

      Number(bill.subtotal || 0),

      bill.discountType || 'amount',
      Number(bill.discount || 0),

      bill.taxLabel || '',
      Number(bill.taxRate || 0),
      Number(bill.taxAmount || 0),

      Number(bill.grandTotal || 0),

      bill.paymentMode || bill.payment || '',
      bill.createdBy || bill.user || '',

      bill.currency || 'INR'
    ]);

  });

  sheet.autoResizeColumns(1, 20);

  return jsonResponse({
    success: true,
    message: 'Bill saved successfully',
    billNo: bill.billNo || bill.id
  });
}


function jsonResponse(data) {

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);

}
