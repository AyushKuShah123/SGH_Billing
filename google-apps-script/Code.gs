const SHEET_ID = '1g-StRRD-sm-UWpyQhqTVodKEdmbUrC4vspKxKQoDiYs';


/**
 * GET
 * Used to test whether the Web App is running.
 */
function doGet() {

  return ContentService
    .createTextOutput(
      JSON.stringify({
        success: true,
        message: 'Shivam Gift House Billing API is running'
      })
    )
    .setMimeType(ContentService.MimeType.JSON);

}


/**
 * POST
 * Receives bill data from the billing website.
 */
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


    /*
     * Expected:
     *
     * {
     *   action: "saveBill",
     *   bill: {...}
     * }
     */


    if (data.action !== 'saveBill') {

      return jsonResponse({
        success: false,
        error: 'Unknown action'
      });

    }


    const bill = data.bill;


    if (!bill) {

      return jsonResponse({
        success: false,
        error: 'Bill data is missing'
      });

    }


    if (!Array.isArray(bill.items)) {

      return jsonResponse({
        success: false,
        error: 'Bill items are missing or invalid'
      });

    }


    if (bill.items.length === 0) {

      return jsonResponse({
        success: false,
        error: 'Bill contains no items'
      });

    }


    return saveBill(bill);


  } catch (error) {

    return jsonResponse({

      success: false,

      error:
        error.toString()

    });

  }

}


/**
 * Save bill to Google Sheet.
 */
function saveBill(bill) {

  if (!bill) {

    return jsonResponse({
      success: false,
      error: 'Bill data is missing'
    });

  }


  const ss =
    SpreadsheetApp.openById(SHEET_ID);


  const date =
    new Date(bill.date || new Date());


  if (isNaN(date.getTime())) {

    return jsonResponse({
      success: false,
      error: 'Invalid bill date'
    });

  }


  const day =
    date.getDate();


  let sheetName;


  if (day <= 15) {

    sheetName =
      Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        'yyyy-MM'
      ) + '-01-15';

  } else {

    sheetName =
      Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        'yyyy-MM'
      ) + '-16-END';

  }


  let sheet =
    ss.getSheetByName(sheetName);


  /**
   * Create the half-month sheet
   * automatically if it doesn't exist.
   */
  if (!sheet) {

    sheet =
      ss.insertSheet(sheetName);


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
      .getRange(1, 1, 1, 20)
      .setFontWeight('bold');


    sheet.setFrozenRows(1);

  }


  const items = Array.isArray(bill.items)
  ? bill.items
  : [];


  /**
   * If there are no items,
   * don't create an empty bill.
   */
  if (items.length === 0) {

    return jsonResponse({

      success: false,
      error: 'Bill contains no items'

    });

  }


  items.forEach(function(item) {


    const quantity =
      Number(
        item.quantity ||
        item.q ||
        0
      );


    const unitPrice =
      Number(
        item.unitPrice ||
        item.u ||
        0
      );


    const itemTotal =
      quantity * unitPrice;


    sheet.appendRow([

      bill.billNo ||
      bill.id ||
      '',


      Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        'dd-MM-yyyy'
      ),


      Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        'HH:mm:ss'
      ),


      bill.customerName ||
      bill.customer ||
      '',


      bill.customerPhone ||
      bill.phone ||
      '',


      bill.customerAddress ||
      bill.address ||
      '',


      item.particular ||
      item.p ||
      '',


      quantity,


      unitPrice,


      itemTotal,


      Number(
        bill.subtotal || 0
      ),


      bill.discountType ||
      'amount',


      Number(
        bill.discount || 0
      ),


      bill.taxLabel ||
      '',


      Number(
        bill.taxRate || 0
      ),


      Number(
        bill.taxAmount || 0
      ),


      Number(
        bill.grandTotal || 0
      ),


      bill.paymentMode ||
      bill.payment ||
      '',


      bill.createdBy ||
      bill.user ||
      '',


      bill.currency ||
      'INR'

    ]);

  });


  sheet.autoResizeColumns(
    1,
    20
  );


  return jsonResponse({

    success: true,

    message:
      'Bill saved successfully',

    billNo:
      bill.billNo ||
      bill.id ||
      ''

  });

}


/**
 * JSON response helper.
 */
function jsonResponse(data) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}
