// ============================================================
// The Juicy Burger — Google Apps Script Backend
// ALL actions via doGet() — zero CORS issues, never redeploy.
// ============================================================
//
// SHEET SETUP:
//   Sheet 1 tab name → "Menu"
//   Sheet 2 tab name → "Orders"
//
// Menu columns:  A:id  B:name  C:category  D:price  E:description  F:image  G:available  H:ingredients
// Order columns: A:id  B:sentAt  C:status  D:paymentStatus  E:paymentMethod  F:paidAt  G:subtotal  H:gst  I:pst  J:total  K:items
// ============================================================

var SS = SpreadsheetApp.getActiveSpreadsheet();

function makeResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SINGLE ROUTER — everything is GET
// ============================================================
function doGet(e) {
  try {
    var p      = e.parameter || {};
    var action = p.action    || '';

    if (action === 'getMenu')        return makeResponse(getMenu());
    if (action === 'setMenuItem')    return makeResponse(setMenuItem(safeParseJson(p.item, null)));
    if (action === 'deleteMenuItem') return makeResponse(deleteMenuItem(p.id));
    if (action === 'getOrders')      return makeResponse(getOrders());
    if (action === 'setOrder')       return makeResponse(setOrder(safeParseJson(p.order, null)));
    if (action === 'updateOrder')    return makeResponse(updateOrder(p.id, safeParseJson(p.fields, {})));
    if (action === 'deleteOrder')    return makeResponse(deleteOrder(p.id));
    if (action === 'seedMenu')       return makeResponse(seedMenuIfEmpty());

    return makeResponse({ error: 'Unknown action: ' + action });
  } catch(err) {
    return makeResponse({ error: err.toString() });
  }
}

function doPost(e) {
  return makeResponse({ error: 'POST disabled — use GET.' });
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
  var rowData = menuRowData(item);

  // Update existing row
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(item.id)) {
      sheet.getRange(i + 1, 1, 1, 8).setValues([rowData]);
      return { ok: true, action: 'updated' };
    }
  }

  // Insert new row
  sheet.appendRow(rowData);
  return { ok: true, action: 'inserted' };
}

function deleteMenuItem(id) {
  if (!id) return { ok: false, error: 'No id provided' };
  var sheet = SS.getSheetByName('Menu');
  if (!sheet) return { ok: false, error: 'Menu sheet not found' };
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {   // reverse loop — safe for deleteRow
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Item not found: ' + id };
}

function menuRowData(item) {
  return [
    String(item.id),
    String(item.name        || ''),
    String(item.category    || ''),
    parseFloat(item.price)  || 0,
    String(item.description || ''),
    String(item.image       || ''),
    item.available !== false,           // default true
    JSON.stringify(item.ingredients || [])
  ];
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
      // Merge only the fields sent
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
    o.sentAt ? new Date(typeof o.sentAt === 'number' ? o.sentAt : Number(o.sentAt)) : '',
    String(o.status        || 'pending'),
    String(o.paymentStatus || 'unpaid'),
    String(o.paymentMethod || ''),
    o.paidAt  ? new Date(typeof o.paidAt  === 'number' ? o.paidAt  : Number(o.paidAt))  : '',
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
// SHEET HELPER — get or create with headers
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
// SEED — only runs if Menu sheet has no data rows
// ============================================================
function seedMenuIfEmpty() {
  var sheet = SS.getSheetByName('Menu');
  if (!sheet) {
    sheet = SS.insertSheet('Menu');
    sheet.getRange(1, 1, 1, 8).setValues(menuHeaders());
  } else {
    var rows = sheet.getDataRange().getValues();
    if (rows.length > 1) return { ok: true, seeded: false };  // already has data
  }

  var defaults = [
    ['item_classic', 'Classic Smash Burger',    'Burgers', 14.99, 'Double smash patty, American cheese, pickles, onion, special sauce',              '', true, '["Double Smash Patty","American Cheese","Pickles","Onion","Special Sauce","Brioche Bun"]'],
    ['item_bbq',     'Smoky BBQ Stack',          'Burgers', 16.99, 'Beef patty, cheddar, crispy onion rings, BBQ sauce, bacon',                       '', true, '["Beef Patty","Cheddar Cheese","Crispy Onion Rings","BBQ Sauce","Bacon","Brioche Bun"]'],
    ['item_truffle', 'Truffle Mushroom Burger',  'Burgers', 18.99, 'Beef patty, sautéed mushrooms, truffle aioli, Swiss cheese, arugula',             '', true, '["Beef Patty","Sautéed Mushrooms","Truffle Aioli","Swiss Cheese","Arugula","Pretzel Bun"]'],
    ['item_spicy',   'Inferno Burger',           'Burgers', 15.99, 'Double patty, jalapeños, pepper jack, sriracha mayo, ghost pepper sauce',          '', true, '["Double Beef Patty","Jalapeños","Pepper Jack Cheese","Sriracha Mayo","Ghost Pepper Sauce","Brioche Bun"]'],
    ['item_veggie',  'Garden Smash',             'Burgers', 13.99, 'Smashed plant patty, vegan cheese, lettuce, tomato, avocado, chipotle mayo',      '', true, '["Plant-Based Patty","Vegan Cheese","Lettuce","Tomato","Avocado","Chipotle Mayo","Brioche Bun"]'],
    ['item_fries',   'Crispy Smash Fries',       'Sides',    5.99, 'Hand-cut fries, seasoned with smash seasoning',                                   '', true, '["Hand-Cut Potatoes","Smash Seasoning","Sea Salt"]'],
    ['item_poutine', 'Juicy Poutine',            'Sides',    8.99, 'Fries, cheese curds, gravy',                                                      '', true, '["Fries","Cheese Curds","Beef Gravy"]'],
    ['item_shake',   'Smash Shake',              'Drinks',   7.99, 'Thick milkshake — vanilla, chocolate, or strawberry',                             '', true, '["Whole Milk Ice Cream","Choice of Flavour","Whipped Cream"]'],
    ['item_soda',    'Fountain Soda',            'Drinks',   3.49, 'Pepsi, Diet Pepsi, 7UP, Orange Crush, Ginger Ale',                                '', true, '["Choice of Soda","Ice"]']
  ];

  sheet.getRange(2, 1, defaults.length, 8).setValues(defaults);

  // Ensure Orders sheet also exists
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
