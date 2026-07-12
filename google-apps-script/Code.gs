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
var EVENTS_SHEET_NAME = 'Events';
var ADMIN_EMAIL = 'xohellobeebee@gmail.com';
var STATUSES = ['New', 'Drawing', 'Ready', 'Picked Up', 'Paid'];
var HEADERS = ['Timestamp', 'Order #', 'Customer', 'Items', 'Total', 'Status', 'Contact', 'Pickup Event', 'Source', 'Notes'];
var EVENT_HEADERS = ['ID', 'Date', 'Name', 'Location', 'Type', 'Published', 'Sort Order', 'Updated At'];
var STATUS_COLORS = {
  'New': '#FFD6D6',
  'Drawing': '#FFE9A0',
  'Ready': '#C5EAD8',
  'Picked Up': '#E6E6E6',
  'Paid': '#C2E4F7'
};

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'events') {
    return json_({ ok: true, events: listEvents_(true) });
  }
  return json_({ ok: true, message: "BeeBee's order and events endpoint is ready" });
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action && /^admin\./.test(data.action)) {
      requireAdmin_(data.credential);
      return handleAdminAction_(data);
    }
    return handleOrder_(data);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function handleOrder_(data) {
  var lock = LockService.getScriptLock();
  try {
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
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function handleAdminAction_(data) {
  if (data.action === 'admin.listEvents') {
    return json_({ ok: true, events: listEvents_(false) });
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (data.action === 'admin.saveEvent') {
      return json_({ ok: true, event: saveEvent_(data.event || {}) });
    }
    if (data.action === 'admin.deleteEvent') {
      deleteEvent_(String(data.id || ''));
      return json_({ ok: true });
    }
    throw new Error('Unknown admin action');
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function requireAdmin_(credential) {
  if (!credential) throw new Error('Sign in required');
  var clientId = PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID');
  if (!clientId) throw new Error('GOOGLE_CLIENT_ID is not configured');

  var response = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(credential),
    { muteHttpExceptions: true }
  );
  if (response.getResponseCode() !== 200) throw new Error('Google sign-in could not be verified');
  var profile = JSON.parse(response.getContentText());
  if (profile.aud !== clientId) throw new Error('Invalid Google client');
  if (String(profile.email_verified) !== 'true') throw new Error('Google email is not verified');
  if (String(profile.email || '').toLowerCase() !== ADMIN_EMAIL) throw new Error('This account is not authorized');
  return profile;
}

function ensureEventsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(EVENTS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(EVENTS_SHEET_NAME);
  if (sheet.getRange(1, 1).getValue() !== EVENT_HEADERS[0]) {
    sheet.clear();
    sheet.getRange(1, 1, 1, EVENT_HEADERS.length).setValues([EVENT_HEADERS])
      .setFontWeight('bold').setBackground('#FFE9A0');
    sheet.setFrozenRows(1);
    sheet.getRange(2, 1, 3, EVENT_HEADERS.length).setValues([
      [Utilities.getUuid(), '2025-07-12', 'Summer Craft Fair', 'Riverside Park, Downtown', 'Market', true, 1, new Date()],
      [Utilities.getUuid(), '2025-07-26', 'Doodle Workshop', 'Local Art Studio - seats limited!', 'Workshop', true, 2, new Date()],
      [Utilities.getUuid(), '2025-08-03', 'Farmers Market Pop-Up', 'Central Square Farmers Market', 'Market', true, 3, new Date()]
    ]);
    [250, 110, 220, 300, 110, 100, 90, 160].forEach(function (width, i) {
      sheet.setColumnWidth(i + 1, width);
    });
  }
  return sheet;
}

function listEvents_(publishedOnly) {
  var sheet = ensureEventsSheet_();
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, EVENT_HEADERS.length).getValues()
    .map(function (row) {
      return {
        id: String(row[0]), date: eventDateValue_(row[1]), name: String(row[2]),
        location: String(row[3]), type: String(row[4] || 'Market'),
        published: row[5] === true, sortOrder: Number(row[6]) || 0
      };
    })
    .filter(function (event) { return !publishedOnly || event.published; })
    .sort(function (a, b) { return a.sortOrder - b.sortOrder || a.date.localeCompare(b.date); });
}

function eventDateValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    // Format in the spreadsheet's own timezone — date cells are stored at midnight
    // spreadsheet-local, so any other zone shifts the date by a day.
    return Utilities.formatDate(value, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  return String(value);
}

function saveEvent_(event) {
  if (!event.name || !event.date || !event.location) throw new Error('Date, name, and location are required');
  var sheet = ensureEventsSheet_();
  var id = String(event.id || Utilities.getUuid());
  var rowValues = [[
    id, String(event.date).slice(0, 20), String(event.name).slice(0, 200),
    String(event.location).slice(0, 300), event.type === 'Workshop' ? 'Workshop' : 'Market',
    event.published === true, Number(event.sortOrder) || 0, new Date()
  ]];
  var targetRow = sheet.getLastRow() + 1;
  if (sheet.getLastRow() >= 2) {
    var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === id) { targetRow = i + 2; break; }
    }
  }
  sheet.getRange(targetRow, 1, 1, EVENT_HEADERS.length).setValues(rowValues);
  return { id: id };
}

function deleteEvent_(id) {
  if (!id) throw new Error('Event ID is required');
  var sheet = ensureEventsSheet_();
  if (sheet.getLastRow() < 2) return;
  var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      sheet.deleteRow(i + 2);
      return;
    }
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
