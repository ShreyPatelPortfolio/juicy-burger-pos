// ============================================================
// The Juicy Burger — Google Apps Script Backend
//
// Menu actions  → doGet()  (GET, avoids CORS issues)
// Order actions → doPost() (POST, handles large payloads)
// ============================================================

var SS = SpreadsheetApp.getActiveSpreadsheet();

function makeResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var p      = e.parameter || {};
    var action = p.action    || '';
    if (action === 'getMenu')        return makeResponse(getMenu());
    if (action === 'setMenuItem')    return makeResponse(setMenuItem(safeParseJson(p.item, null)));
    if (action === 'deleteMenuItem') return makeResponse(deleteMenuItem(p.id));
    if (action === 'getOrders')      return makeResponse(getOrders());
    if (action === 'seedMenu')       return makeResponse(seedMenuIfEmpty());
    return makeResponse({ error: 'Unknown GET action: ' + action });
  } catch(err) {
    return makeResponse({ error: err.toString() });
  }
}

function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var action = data.action || '';
    if (action === 'setOrder')    return makeResponse(setOrder(data.order));
    if (action === 'updateOrder') return makeResponse(updateOrder(data.id, data.fields));
    if (action === 'deleteOrder') return makeResponse(deleteOrder(data.id));
    return makeResponse({ error: 'Unknown POST action: ' + action });
  } catch(err) {
    return makeResponse({ error: err.toString() });
  }
}

// ============================================================
// MENU
// ============================================================
function getMenu() {
  var sheet = getOrCreateSheet('Menu', menuHeaders());
  var rows  = sheet.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    var id = String(r[0]);
    result[id] = {
      id:          id,
      name:        String(r[1] || ''),
      category:    String(r[2] || ''),
      price:       parseFloat(r[3]) || 0,
      description: String(r[4] || ''),
      image:       String(r[5] || ''),
      available:   r[6] === true || r[6] === 'TRUE' || r[6] === 1 || r[6] === 'true',
      ingredients: safeParseJson(r[7], [])
    };
  }
  return result;
}

function setMenuItem(item) {
  if (!item || !item.id) return { ok: false, error: 'Invalid item — missing id' };
  var sheet   = getOrCreateSheet('Menu', menuHeaders());
  var rows    = sheet.getDataRange().getValues();
  var rowData = [
    String(item.id),
    String(item.name        || ''),
    String(item.category    || ''),
    parseFloat(item.price)  || 0,
    String(item.description || ''),
    String(item.image       || ''),
    item.available !== false,
    JSON.stringify(item.ingredients || [])
  ];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(item.id)) {
      sheet.getRange(i + 1, 1, 1, 8).setValues([rowData]);
      return { ok: true, action: 'updated' };
    }
  }
  sheet.appendRow(rowData);
  return { ok: true, action: 'inserted' };
}

function deleteMenuItem(id) {
  if (!id) return { ok: false, error: 'No id provided' };
  var sheet = SS.getSheetByName('Menu');
  if (!sheet) return { ok: false, error: 'Menu sheet not found' };
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Item not found: ' + id };
}

function menuHeaders() {
  return [['id','name','category','price','description','image','available','ingredients']];
}

// ============================================================
// ORDERS
// ============================================================
function getOrders() {
  var sheet = getOrCreateSheet('Orders', orderHeaders());
  var rows  = sheet.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    var id = String(r[0]);
    result[id] = {
      id:            id,
      sentAt:        r[1] ? new Date(r[1]).getTime() : 0,
      status:        String(r[2] || 'pending'),
      paymentStatus: String(r[3] || 'unpaid'),
      paymentMethod: String(r[4] || ''),
      paidAt:        r[5] ? new Date(r[5]).getTime() : null,
      subtotal:      parseFloat(r[6]) || 0,
      gst:           parseFloat(r[7]) || 0,
      pst:           parseFloat(r[8]) || 0,
      total:         parseFloat(r[9]) || 0,
      items:         safeParseJson(r[10], [])
    };
  }
  return result;
}

function setOrder(order) {
  if (!order || !order.id) return { ok: false, error: 'Invalid order — missing id' };
  var sheet   = getOrCreateSheet('Orders', orderHeaders());
  var rows    = sheet.getDataRange().getValues();
  var rowData = orderToRow(order);
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(order.id)) {
      sheet.getRange(i + 1, 1, 1, 11).setValues([rowData]);
      return { ok: true, action: 'updated' };
    }
  }
  sheet.appendRow(rowData);
  return { ok: true, action: 'inserted' };
}

function updateOrder(id, fields) {
  if (!id) return { ok: false, error: 'No id provided' };
  var sheet = SS.getSheetByName('Orders');
  if (!sheet) return { ok: false, error: 'Orders sheet not found' };
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      var current = {
        id:            rows[i][0],
        sentAt:        rows[i][1],
        status:        rows[i][2],
        paymentStatus: rows[i][3],
        paymentMethod: rows[i][4],
        paidAt:        rows[i][5],
        subtotal:      rows[i][6],
        gst:           rows[i][7],
        pst:           rows[i][8],
        total:         rows[i][9],
        items:         safeParseJson(rows[i][10], [])
      };
      if (fields.status        !== undefined) current.status        = fields.status;
      if (fields.paymentStatus !== undefined) current.paymentStatus = fields.paymentStatus;
      if (fields.paymentMethod !== undefined) current.paymentMethod = fields.paymentMethod;
      if (fields.paidAt        !== undefined) current.paidAt        = fields.paidAt ? new Date(fields.paidAt) : null;
      sheet.getRange(i + 1, 1, 1, 11).setValues([orderToRow(current)]);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Order not found: ' + id };
}

function deleteOrder(id) {
  if (!id) return { ok: false, error: 'No id provided' };
  var sheet = SS.getSheetByName('Orders');
  if (!sheet) return { ok: false, error: 'Orders sheet not found' };
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Order not found: ' + id };
}

function orderToRow(o) {
  return [
    String(o.id),
    o.sentAt  ? new Date(typeof o.sentAt  === 'number' ? o.sentAt  : Number(o.sentAt)) : '',
    String(o.status        || 'pending'),
    String(o.paymentStatus || 'unpaid'),
    String(o.paymentMethod || ''),
    o.paidAt  ? new Date(typeof o.paidAt  === 'number' ? o.paidAt  : Number(o.paidAt)) : '',
    parseFloat(o.subtotal) || 0,
    parseFloat(o.gst)      || 0,
    parseFloat(o.pst)      || 0,
    parseFloat(o.total)    || 0,
    JSON.stringify(o.items || [])
  ];
}

function orderHeaders() {
  return [['id','sentAt','status','paymentStatus','paymentMethod','paidAt','subtotal','gst','pst','total','items']];
}

// ============================================================
// SHEET HELPER
// ============================================================
function getOrCreateSheet(name, headers) {
  var sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  }
  return sheet;
}

// ============================================================
// SEED — your exact menu with readable name-based IDs
// Only runs if Menu sheet has zero data rows
// ============================================================
function seedMenuIfEmpty() {
  var sheet = SS.getSheetByName('Menu');
  if (!sheet) {
    sheet = SS.insertSheet('Menu');
    sheet.getRange(1, 1, 1, 8).setValues(menuHeaders());
  } else {
    var rows = sheet.getDataRange().getValues();
    if (rows.length > 1) return { ok: true, seeded: false };
  }

  var defaults = [
    ['item_classic_smash_burger', 'Classic Smash Burger', 'Burgers', 9.99, 'Double smash patty, American cheese, pickles, onion, special sauce', '', true, '["Double Smash Patty","American Cheese","Pickles","Onion","Special Sauce","Brioche Bun"]'],
    ['item_crispy_smash_fries', 'Crispy Smash Fries', 'Sides', 4.99, 'Hand-cut fries, seasoned with smash seasoning', '', true, '["Hand-Cut Potatoes","Smash Seasoning","Sea Salt"]'],
    ['item_juicy_poutine', 'Juicy Poutine', 'Sides', 7.99, 'Fries, cheese curds, gravy', '', true, '["Fries","Cheese Curds","Beef Gravy"]'],
    ['item_quarter_pound_burger', 'Quarter Pound Burger', 'Burgers', 11.99, 'Quarter Pound Beef Patty, Pickles, Onions, Lettuce, Tomatoes, Swiss Cheese', '', true, '["Quarter Pound Beef patty","Pickles","Onions","Tomatoes","Lettuce","Swiss Cheese"]'],
    ['item_12_chicken_nuggets', '12 Chicken Nuggets', 'Chicken Nuggets', 9.99, 'Chicken Nuggets with side dressing of BBQ Sauce', '', true, '["Chicken Nuggets","BBQ Sauce"]'],
    ['item_combo_1', 'Combo 1', 'Specials', 14.99, 'Smash Burger/Nuggets + Fries + Any Fruity Tea (Regular size)', '', true, '[]'],
    ['item_combo_2', 'Combo 2', 'Specials', 16.99, 'Quarter Pound Burger + Fries + Any Fruity Tea (Regular size)', '', true, '[]'],
    ['item_extra_veggie', 'Extra Veggie', 'Extra', 0.50, '', '', true, '["Onion","Tomatoes","Lettuce","Pickles"]'],
    ['item_extra_cheese', 'Extra Cheese', 'Extra', 0.80, '', '', true, '["Swiss cheese"]']
  ];

  sheet.getRange(2, 1, defaults.length, 8).setValues(defaults);
  getOrCreateSheet('Orders', orderHeaders());
  return { ok: true, seeded: true };
}

// ============================================================
// UTILITY
// ============================================================
function safeParseJson(str, fallback) {
  if (str === null || str === undefined || str === '') return fallback;
  try { return JSON.parse(str); } catch(e) { return fallback; }
}
