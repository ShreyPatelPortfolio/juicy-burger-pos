// ============================================================
// The Juicy Burger — Google Apps Script Backend
// Paste this entire file into your Apps Script editor.
// ============================================================
//
// SHEET SETUP:
//   Sheet 1 tab name → "Menu"
//   Sheet 2 tab name → "Orders"
//
// Menu sheet columns (row 1 = headers):
//   A: id | B: name | C: category | D: price | E: description
//   F: image | G: available | H: ingredients (JSON array string)
//
// Orders sheet columns (row 1 = headers):
//   A: id | B: sentAt | C: status | D: paymentStatus | E: paymentMethod
//   F: paidAt | G: subtotal | H: gst | I: pst | J: total
//   K: items (JSON string)
// ============================================================

var SS = SpreadsheetApp.getActiveSpreadsheet();

// ---- CORS helper ----
function makeResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- Router ----
function doGet(e) {
  try {
    var action = e.parameter.action;
    if (action === 'getMenu')   return makeResponse(getMenu());
    if (action === 'getOrders') return makeResponse(getOrders());
    return makeResponse({ error: 'Unknown action' });
  } catch(err) {
    return makeResponse({ error: err.toString() });
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;

    if (action === 'setMenuItem')    return makeResponse(setMenuItem(data.item));
    if (action === 'deleteMenuItem') return makeResponse(deleteMenuItem(data.id));
    if (action === 'setOrder')       return makeResponse(setOrder(data.order));
    if (action === 'updateOrder')    return makeResponse(updateOrder(data.id, data.fields));
    if (action === 'deleteOrder')    return makeResponse(deleteOrder(data.id));
    if (action === 'seedMenu')       return makeResponse(seedMenuIfEmpty());
    return makeResponse({ error: 'Unknown action' });
  } catch(err) {
    return makeResponse({ error: err.toString() });
  }
}

// ============================================================
// MENU
// ============================================================
function getMenu() {
  var sheet = SS.getSheetByName('Menu');
  if (!sheet) return {};
  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return {};
  var result = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue; // skip blank rows
    var id = String(r[0]);
    result[id] = {
      id:          id,
      name:        r[1] || '',
      category:    r[2] || '',
      price:       parseFloat(r[3]) || 0,
      description: r[4] || '',
      image:       r[5] || '',
      available:   r[6] === true || r[6] === 'TRUE' || r[6] === 1,
      ingredients: safeParseJson(r[7], [])
    };
  }
  return result;
}

function setMenuItem(item) {
  var sheet = SS.getSheetByName('Menu');
  if (!sheet) { SS.insertSheet('Menu'); sheet = SS.getSheetByName('Menu'); ensureMenuHeaders(sheet); }
  var rows = sheet.getDataRange().getValues();
  // Find existing row by id
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(item.id)) {
      sheet.getRange(i + 1, 1, 1, 8).setValues([[
        item.id, item.name, item.category, item.price,
        item.description || '', item.image || '',
        item.available !== false,
        JSON.stringify(item.ingredients || [])
      ]]);
      return { ok: true, action: 'updated' };
    }
  }
  // Append new
  sheet.appendRow([
    item.id, item.name, item.category, item.price,
    item.description || '', item.image || '',
    item.available !== false,
    JSON.stringify(item.ingredients || [])
  ]);
  return { ok: true, action: 'inserted' };
}

function deleteMenuItem(id) {
  var sheet = SS.getSheetByName('Menu');
  if (!sheet) return { ok: false };
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Not found' };
}

function ensureMenuHeaders(sheet) {
  sheet.getRange(1, 1, 1, 8).setValues([[
    'id','name','category','price','description','image','available','ingredients'
  ]]);
}

// ============================================================
// ORDERS
// ============================================================
function getOrders() {
  var sheet = SS.getSheetByName('Orders');
  if (!sheet) return {};
  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return {};
  var result = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    var id = String(r[0]);
    result[id] = {
      id:            id,
      sentAt:        r[1] ? new Date(r[1]).getTime() : 0,
      status:        r[2] || 'pending',
      paymentStatus: r[3] || 'unpaid',
      paymentMethod: r[4] || '',
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
  var sheet = SS.getSheetByName('Orders');
  if (!sheet) { SS.insertSheet('Orders'); sheet = SS.getSheetByName('Orders'); ensureOrderHeaders(sheet); }
  var rows = sheet.getDataRange().getValues();
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
  var sheet = SS.getSheetByName('Orders');
  if (!sheet) return { ok: false };
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      // Read current, merge fields, write back
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
      // Apply updates
      if (fields.status        !== undefined) current.status        = fields.status;
      if (fields.paymentStatus !== undefined) current.paymentStatus = fields.paymentStatus;
      if (fields.paymentMethod !== undefined) current.paymentMethod = fields.paymentMethod;
      if (fields.paidAt        !== undefined) current.paidAt        = fields.paidAt ? new Date(fields.paidAt) : null;
      sheet.getRange(i + 1, 1, 1, 11).setValues([orderToRow(current)]);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Order not found' };
}

function deleteOrder(id) {
  var sheet = SS.getSheetByName('Orders');
  if (!sheet) return { ok: false };
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false };
}

function orderToRow(o) {
  return [
    o.id,
    o.sentAt ? new Date(typeof o.sentAt === 'number' ? o.sentAt : o.sentAt) : '',
    o.status        || 'pending',
    o.paymentStatus || 'unpaid',
    o.paymentMethod || '',
    o.paidAt  ? new Date(typeof o.paidAt  === 'number' ? o.paidAt  : o.paidAt)  : '',
    o.subtotal || 0,
    o.gst      || 0,
    o.pst      || 0,
    o.total    || 0,
    JSON.stringify(o.items || [])
  ];
}

function ensureOrderHeaders(sheet) {
  sheet.getRange(1, 1, 1, 11).setValues([[
    'id','sentAt','status','paymentStatus','paymentMethod',
    'paidAt','subtotal','gst','pst','total','items'
  ]]);
}

// ============================================================
// SEED default menu if Menu sheet is empty
// ============================================================
function seedMenuIfEmpty() {
  var sheet = SS.getSheetByName('Menu');
  if (!sheet) {
    sheet = SS.insertSheet('Menu');
    ensureMenuHeaders(sheet);
  } else {
    var rows = sheet.getDataRange().getValues();
    if (rows.length > 1) return { ok: true, seeded: false }; // already has data
  }

  var defaults = [
    ['item_classic','Classic Smash Burger','Burgers',14.99,'Double smash patty, American cheese, pickles, onion, special sauce','',true,'["Double Smash Patty","American Cheese","Pickles","Onion","Special Sauce","Brioche Bun"]'],
    ['item_bbq','Smoky BBQ Stack','Burgers',16.99,'Beef patty, cheddar, crispy onion rings, BBQ sauce, bacon','',true,'["Beef Patty","Cheddar Cheese","Crispy Onion Rings","BBQ Sauce","Bacon","Brioche Bun"]'],
    ['item_truffle','Truffle Mushroom Burger','Burgers',18.99,'Beef patty, sautéed mushrooms, truffle aioli, Swiss cheese, arugula','',true,'["Beef Patty","Sautéed Mushrooms","Truffle Aioli","Swiss Cheese","Arugula","Pretzel Bun"]'],
    ['item_spicy','Inferno Burger','Burgers',15.99,'Double patty, jalapeños, pepper jack, sriracha mayo, ghost pepper sauce','',true,'["Double Beef Patty","Jalapeños","Pepper Jack Cheese","Sriracha Mayo","Ghost Pepper Sauce","Brioche Bun"]'],
    ['item_veggie','Garden Smash','Burgers',13.99,'Smashed plant patty, vegan cheese, lettuce, tomato, avocado, chipotle mayo','',true,'["Plant-Based Patty","Vegan Cheese","Lettuce","Tomato","Avocado","Chipotle Mayo","Brioche Bun"]'],
    ['item_fries','Crispy Smash Fries','Sides',5.99,'Hand-cut fries, seasoned with smash seasoning','',true,'["Hand-Cut Potatoes","Smash Seasoning","Sea Salt"]'],
    ['item_poutine','Juicy Poutine','Sides',8.99,'Fries, cheese curds, gravy','',true,'["Fries","Cheese Curds","Beef Gravy"]'],
    ['item_shake','Smash Shake','Drinks',7.99,'Thick milkshake — vanilla, chocolate, or strawberry','',true,'["Whole Milk Ice Cream","Choice of Flavour","Whipped Cream"]'],
    ['item_soda','Fountain Soda','Drinks',3.49,'Pepsi, Diet Pepsi, 7UP, Orange Crush, Ginger Ale','',true,'["Choice of Soda","Ice"]']
  ];

  // Write headers + data
  sheet.getRange(1, 1, 1, 8).setValues([['id','name','category','price','description','image','available','ingredients']]);
  sheet.getRange(2, 1, defaults.length, 8).setValues(defaults);

  // Ensure Orders sheet exists too
  if (!SS.getSheetByName('Orders')) {
    var os = SS.insertSheet('Orders');
    ensureOrderHeaders(os);
  }

  return { ok: true, seeded: true };
}

// ============================================================
// Utility
// ============================================================
function safeParseJson(str, fallback) {
  try { return JSON.parse(str); } catch(e) { return fallback; }
}
