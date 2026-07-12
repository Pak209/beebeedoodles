/**
 * BeeBee's Doodle Booth — order intake backend.
 *
 * Lives inside the "BeeBee Orders" Google Sheet (Extensions → Apps Script).
 * Deployed as a Web App (Execute as: Me / Access: Anyone), it receives order
 * JSON from beebeedoodles.com and the booth kiosk, appends one row per order,
 * and returns an order number the site shows to the customer.
 *
 * The sheet self-initializes: headers, status dropdown, and colors are
 * created on the first order, so no manual setup run is required.
 */

var SHEET_NAME = 'Orders';
var STATUSES = ['New', 'Drawing', 'Ready', 'Picked Up', 'Paid'];
var HEADERS = ['Timestamp', 'Order #', 'Customer', 'Items', 'Total', 'Status', 'Contact', 'Pickup Event', 'Source', 'Notes'];
var STATUS_COLORS = {
  'New': '#FFD6D6',
  'Drawing': '#FFE9A0',
  'Ready': '#C5EAD8',
  'Picked Up': '#E6E6E6',
  'Paid': '#C2E4F7'
};

function doGet() {
  return json_({ ok: true, message: "BeeBee's order endpoint is buzzing 🐝" });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var data = JSON.parse(e.postData.contents);
    if (!data.name || !data.contact) {
      return json_({ ok: false, error: 'Missing name or contact' });
    }
    if (!data.items || !data.items.length) {
      return json_({ ok: false, error: 'Order has no items' });
    }

    lock.waitLock(15000);
    var sheet = ensureSheet_();
    var orderNumber = nextOrderNumber_();
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy h:mm a');

    sheet.appendRow([
      timestamp,
      orderNumber,
      String(data.name).slice(0, 200),
      formatItems_(data.items),
      Number(data.total) || 0,
      'New',
      String(data.contact).slice(0, 200),
      String(data.pickupEvent || '').slice(0, 200),
      String(data.source || 'website').slice(0, 50),
      String(data.notes || '').slice(0, 2000)
    ]);

    return json_({ ok: true, orderNumber: orderNumber });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

/** One readable line per item, add-on selections in parentheses. */
function formatItems_(items) {
  return items.map(function (item, i) {
    var opts = (item.options || [])
      .filter(function (o) {
        return o.value && o.value !== 'Not selected' && !/^No /.test(o.value);
      })
      .map(function (o) { return o.label + ': ' + o.value; })
      .join(' · ');
    var line = (i + 1) + '. ' + item.name + ' — $' + (Number(item.price) || 0);
    return opts ? line + ' (' + opts + ')' : line;
  }).join('\n').slice(0, 5000);
}

/** Sequential BB-#### numbers that survive row deletions. */
function nextOrderNumber_() {
  var props = PropertiesService.getScriptProperties();
  var next = Number(props.getProperty('lastOrderNumber') || 1000) + 1;
  props.setProperty('lastOrderNumber', String(next));
  return 'BB-' + next;
}

function ensureSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME, 0);
  }
  if (sheet.getRange(1, 1).getValue() === HEADERS[0]) {
    return sheet;
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold')
    .setBackground('#DDD0F5');
  sheet.setFrozenRows(1);

  var widths = [150, 90, 150, 380, 70, 110, 170, 170, 90, 260];
  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });

  var statusRange = sheet.getRange(2, 6, sheet.getMaxRows() - 1, 1);
  statusRange.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build()
  );

  var rules = STATUSES.map(function (status) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(status)
      .setBackground(STATUS_COLORS[status])
      .setRanges([statusRange])
      .build();
  });
  sheet.setConditionalFormatRules(rules);

  sheet.getRange(2, 4, sheet.getMaxRows() - 1, 1).setWrap(true);
  sheet.getRange(2, 10, sheet.getMaxRows() - 1, 1).setWrap(true);
  sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1).setNumberFormat('$#,##0');

  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
