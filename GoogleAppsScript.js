/**
 * ZALVRA Store — Google Apps Script (Database)
 * ============================================
 *
 * SETUP:
 * 1. Open Google Sheet → Extensions → Apps Script
 * 2. Paste this file → Save
 * 3. Run function "setup" once → Allow permissions
 * 4. Deploy → New deployment → Web app
 *    Execute as: Me | Who has access: Anyone
 * 5. Paste Web App URL in Admin
 */

var PRODUCTS_SHEET = 'Products';
var ORDERS_SHEET = 'Orders';
var PRODUCT_HEADERS = ['id', 'name', 'price', 'oldPrice', 'cat', 'emoji', 'rating', 'desc', 'images'];
var ORDER_HEADERS = ['id', 'date', 'name', 'email', 'phone', 'address', 'payment', 'product', 'emoji', 'qty', 'total', 'status'];
var CACHE_TTL = 8;

function setup() {
  setupSheets_();
  clearCache_();
  return 'ZALVRA setup OK';
}

function doGet(e) {
  try {
    var data = parseRequest_(e);
    return handleAction_(data.action || 'ping', data);
  } catch (err) {
    return json_({ error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    var data = parseRequest_(e);
    return handleAction_(data.action || 'ping', data);
  } catch (err) {
    return json_({ error: String(err.message || err) });
  }
}

/** Accept JSON body OR form fields OR ?action=&data= */
function parseRequest_(e) {
  var data = {};
  e = e || {};

  // URL / form parameters
  if (e.parameter) {
    if (e.parameter.action) data.action = e.parameter.action;
    if (e.parameter.data) {
      try {
        var parsed = JSON.parse(e.parameter.data);
        for (var k in parsed) {
          if (parsed.hasOwnProperty(k)) data[k] = parsed[k];
        }
      } catch (ignore) {}
    }
    // also allow direct product/order JSON strings
    if (e.parameter.product) {
      try { data.product = JSON.parse(e.parameter.product); } catch (ignore) { data.product = e.parameter.product; }
    }
    if (e.parameter.order) {
      try { data.order = JSON.parse(e.parameter.order); } catch (ignore) { data.order = e.parameter.order; }
    }
    if (e.parameter.orderId) data.orderId = e.parameter.orderId;
    if (e.parameter.productId) data.productId = e.parameter.productId;
    if (e.parameter.status) data.status = e.parameter.status;
    // Flat order fields (reliable for GET — avoids nested JSON truncation)
    var flatKeys = ['id','date','name','email','phone','city','address','payment','product','emoji','qty','total'];
    var hasFlat = false;
    for (var fi = 0; fi < flatKeys.length; fi++) {
      var fk = flatKeys[fi];
      if (e.parameter[fk] !== undefined && e.parameter[fk] !== '') {
        hasFlat = true;
        break;
      }
    }
    if (hasFlat && !data.order) {
      data.order = {};
      for (var fj = 0; fj < flatKeys.length; fj++) {
        var key = flatKeys[fj];
        if (e.parameter[key] !== undefined) data.order[key] = e.parameter[key];
      }
    }
  }

  // Raw JSON body (text/plain or application/json)
  if (e.postData && e.postData.contents) {
    var ctype = String(e.postData.type || '');
    if (ctype.indexOf('application/x-www-form-urlencoded') === -1) {
      try {
        var body = JSON.parse(e.postData.contents);
        for (var key in body) {
          if (body.hasOwnProperty(key)) data[key] = body[key];
        }
      } catch (ignore) {}
    }
  }

  return data;
}

function handleAction_(action, data) {
  setupSheets_();

  switch (action) {
    case 'ping':
    case 'setup':
      return json_({
        ok: true,
        message: 'ZALVRA API connected',
        sheets: { products: PRODUCTS_SHEET, orders: ORDERS_SHEET }
      });

    case 'getProducts':
      return json_({ products: getProductsFast_() });

    case 'addProduct':
      addProduct_(data.product);
      clearCacheKey_('products');
      return json_({ ok: true });

    case 'updateProduct':
      updateProduct_(data.product);
      clearCacheKey_('products');
      return json_({ ok: true });

    case 'deleteProduct':
      deleteById_(getProductsSheet_(), data.productId);
      clearCacheKey_('products');
      return json_({ ok: true });

    case 'getOrders':
      return json_({ orders: getOrdersFast_() });

    case 'addOrder':
      var ord = data.order;
      if (!ord || (!ord.name && !ord.phone)) {
        // fallback: treat top-level fields as the order
        ord = data;
      }
      addOrder_(ord);
      clearCacheKey_('orders');
      return json_({ ok: true });

    case 'updateOrder':
      updateOrderStatus_(data.orderId, data.status);
      clearCacheKey_('orders');
      return json_({ ok: true });

    case 'deleteOrder':
      deleteById_(getOrdersSheet_(), data.orderId);
      clearCacheKey_('orders');
      return json_({ ok: true });

    default:
      return json_({ error: 'Unknown action: ' + action });
  }
}

function getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet. Open script from Sheet: Extensions → Apps Script');
  }
  return ss;
}

function setupSheets_() {
  var ss = getSpreadsheet_();
  getOrCreateSheet_(ss, PRODUCTS_SHEET, PRODUCT_HEADERS);
  getOrCreateSheet_(ss, ORDERS_SHEET, ORDER_HEADERS);
}

function getOrCreateSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  var firstCell = '';
  try { firstCell = String(sheet.getRange(1, 1).getValue() || ''); } catch (ignore) {}
  if (firstCell.toLowerCase() !== String(headers[0]).toLowerCase()) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    try {
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#755139').setFontColor('#ffffff');
    } catch (ignore) {}
  }
  return sheet;
}

function getProductsSheet_() {
  return getOrCreateSheet_(getSpreadsheet_(), PRODUCTS_SHEET, PRODUCT_HEADERS);
}

function getOrdersSheet_() {
  return getOrCreateSheet_(getSpreadsheet_(), ORDERS_SHEET, ORDER_HEADERS);
}

function cacheGet_(key) {
  try {
    var raw = CacheService.getScriptCache().get('z_' + key);
    if (raw) return JSON.parse(raw);
  } catch (ignore) {}
  return null;
}

function cachePut_(key, value) {
  try { CacheService.getScriptCache().put('z_' + key, JSON.stringify(value), CACHE_TTL); } catch (ignore) {}
}

function clearCacheKey_(key) {
  try { CacheService.getScriptCache().remove('z_' + key); } catch (ignore) {}
}

function clearCache_() {
  clearCacheKey_('products');
  clearCacheKey_('orders');
}

function getProductsFast_() {
  var cached = cacheGet_('products');
  if (cached) return cached;
  var list = readProducts_();
  cachePut_('products', list);
  return list;
}

function readProducts_() {
  var sheet = getProductsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var numCols = Math.max(sheet.getLastColumn(), PRODUCT_HEADERS.length);
  var values = sheet.getRange(1, 1, lastRow, numCols).getValues();
  var headers = values[0];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (row[0] === '' || row[0] === null || row[0] === undefined) continue;
    var obj = {};
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i];
      if (!h) continue;
      var val = row[i];
      if (h === 'id' || h === 'price' || h === 'rating') val = Number(val) || 0;
      else if (h === 'oldPrice') val = (val === '' || val === null || val === undefined) ? null : Number(val);
      else if (h === 'images') val = imagesFromString_(val);
      obj[h] = val;
    }
    if (obj.id) out.push(obj);
  }
  return out;
}

function productRow_(p) {
  p = p || {};
  return [
    p.id, p.name || '', p.price || 0,
    (p.oldPrice === null || p.oldPrice === undefined || p.oldPrice === '') ? '' : p.oldPrice,
    p.cat || '', p.emoji || '', p.rating || 4.5, p.desc || '', imagesToString_(p.images)
  ];
}

function addProduct_(product) { getProductsSheet_().appendRow(productRow_(product)); }

function updateProduct_(product) {
  if (!product) return;
  var sheet = getProductsSheet_();
  var row = findRowById_(sheet, product.id);
  if (row > 0) sheet.getRange(row, 1, 1, 9).setValues([productRow_(product)]);
  else addProduct_(product);
}

function getOrdersFast_() {
  // Always fresh — orders must show in admin immediately
  var list = readOrders_();
  return list;
}

function readOrders_() {
  var sheet = getOrdersSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var numCols = Math.max(sheet.getLastColumn(), ORDER_HEADERS.length);
  var values = sheet.getRange(1, 1, lastRow, numCols).getValues();
  var headers = values[0];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row[0]) continue;
    var obj = {};
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i];
      if (!h) continue;
      var val = row[i];
      if (h === 'qty' || h === 'total') val = Number(val) || 0;
      obj[h] = val;
    }
    if (obj.id) out.push(obj);
  }
  out.reverse();
  return out;
}

function addOrder_(order) {
  order = order || {};
  // Reject empty / incomplete orders so blank rows are not created
  if (!order.name && !order.phone && !order.product) {
    throw new Error('Incomplete order: name, phone, and product required');
  }
  var id = order.id || ('ORD-' + new Date().getTime());
  var qty = Number(order.qty) || 1;
  var total = Number(order.total) || 0;
  getOrdersSheet_().appendRow([
    id,
    order.date || new Date().toLocaleString(),
    order.name || '',
    order.email || '',
    order.phone || '',
    order.address || '',
    order.payment || 'cod',
    order.product || '',
    order.emoji || '',
    qty,
    total,
    order.status || 'new'
  ]);
}

function updateOrderStatus_(orderId, status) {
  var row = findRowById_(getOrdersSheet_(), orderId);
  if (row > 0) getOrdersSheet_().getRange(row, 12).setValue(status);
}

function findRowById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var ids = sheet.getRange(2, 1, lastRow, 1).getValues();
  var target = String(id);
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === target) return i + 2;
  }
  return 0;
}

function deleteById_(sheet, id) {
  var row = findRowById_(sheet, id);
  if (row > 0) sheet.deleteRow(row);
}

function imagesToString_(images) {
  if (!images) return '';
  if (typeof images === 'object' && images.length >= 0) {
    var parts = [];
    for (var i = 0; i < images.length; i++) if (images[i]) parts.push(String(images[i]));
    return parts.join('|');
  }
  return String(images);
}

function imagesFromString_(val) {
  if (val === null || val === undefined || val === '') return [];
  if (typeof val === 'object' && val.length >= 0) return val;
  var parts = String(val).split('|');
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var s = parts[i].replace(/^\s+|\s+$/g, '');
    if (s) out.push(s);
  }
  return out;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
