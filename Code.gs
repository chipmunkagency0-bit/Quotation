// ================================================================
// ME QUOTATION SYSTEM - Apps Script Backend
// Deploy as: Web App -> Execute as Me -> Anyone
// ================================================================

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  var ss = getSS();
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  var configs = {
    Users:      ['id', 'password', 'role'],
    Items:      ['id', 'name', 'code', 'price', 'brand', 'folder'],
    Quotations: ['id', 'date', 'client', 'qno', 'site', 'rows_json', 'total_before', 'total_tax', 'total_after'],
    Folders:    ['folder_name'],
    Config:     ['key', 'value'],
    Devices:    ['token', 'label', 'approved', 'created', 'last_seen']
  };
  if (configs[name]) sh.appendRow(configs[name]);
  if (name === 'Users')  sh.appendRow(['admin', 'admin123', 'admin']);
  if (name === 'Config') sh.appendRow(['gem_enc', '']);
  return sh;
}

function sheetRows(name) {
  var sh = getSheet(name);
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  var hdrs = vals[0];
  return vals.slice(1).map(function(row) {
    var o = {};
    hdrs.forEach(function(h, i) { o[h] = (row[i] === '' || row[i] === null) ? null : row[i]; });
    return o;
  });
}

//  RESPOND - supports both JSON and JSONP 
function respond(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

//  ROUTER 
function doGet(e) {
  var p  = (e && e.parameter) ? e.parameter : {};
  var cb = p.callback || null;   // JSONP callback name
  var a  = p.action || '';

  try {
    var result;
    if      (a === 'ping')    result = { ok: true, msg: 'API running' };
    else if (a === 'setup')   result = setup();
    else if (a === 'users')   result = getUsers();
    else if (a === 'login')   result = loginUser(p.id, p.password);
    else if (a === 'items')   result = getItems();
    else if (a === 'folders') result = getFolders();
    else if (a === 'history') result = getHistory();
    else if (a === 'config')  result = getConfig();
    // Write operations via GET params (JSONP compatible)
    else if (a === 'save_items')       result = saveItems(safeJSON(p.items, []), p.clear === 'true');
    else if (a === 'append_items')     result = appendItems(safeJSON(p.items, []));
    else if (a === 'delete_item')      result = deleteItemById(p.id);
    else if (a === 'save_folders')     result = saveFolders(safeJSON(p.folders, []));
    else if (a === 'save_quotation')   result = saveQuotation(safeJSON(p.quotation, null));
    else if (a === 'delete_quotation') result = deleteQuotation(p.id);
    else if (a === 'clear_history')    result = clearHistory();
    else if (a === 'register_device')  result = registerDevice(p.token, p.label);
    else result = { ok: true, msg: 'ME Quotation API' };

    return respond(result, cb);
  } catch(err) {
    return respond({ ok: false, error: err.message }, cb);
  }
}

// Keep doPost as fallback
function doPost(e) {
  var b = {};
  try { b = JSON.parse(e.postData.contents); } catch(x) {}
  var a = b.action || '';
  try {
    var result;
    if      (a === 'save_items')       result = saveItems(b.items, b.clear);
    else if (a === 'append_items')     result = appendItems(b.items);
    else if (a === 'delete_item')      result = deleteItemById(b.id);
    else if (a === 'save_folders')     result = saveFolders(b.folders);
    else if (a === 'save_quotation')   result = saveQuotation(b.quotation);
    else if (a === 'delete_quotation') result = deleteQuotation(b.id);
    else if (a === 'clear_history')    result = clearHistory();
    else if (a === 'register_device')  result = registerDevice(b.token, b.label);
    else result = { ok: false, error: 'unknown: ' + a };
    return respond(result, null);
  } catch(err) {
    return respond({ ok: false, error: err.message }, null);
  }
}

//  SETUP 
function setup() {
  ['Users','Items','Quotations','Folders','Config','Devices'].forEach(getSheet);
  return { ok: true, msg: 'All sheets ready' };
}

//  USERS 
function getUsers() {
  var rows = sheetRows('Users').filter(function(r) { return r.id && r.password; });
  return { ok: true, users: rows.map(function(r) {
    return {
      id:       String(r.id       || '').trim(),
      role:     String(r.role     || 'admin').trim()
    };
  })};
}

function loginUser(id, password) {
  var inId = String(id || '').trim().toLowerCase();
  var inPw = String(password || '').trim();
  if (!inId || !inPw) return { ok: false, error: 'missing credentials' };

  var rows = sheetRows('Users').filter(function(r) { return r.id && r.password; });
  for (var i = 0; i < rows.length; i++) {
    var rId = String(rows[i].id || '').trim().toLowerCase();
    var rPw = String(rows[i].password || '').trim();
    if (rId === inId && rPw === inPw) {
      return { ok: true, user: { id: String(rows[i].id || '').trim(), role: String(rows[i].role || 'admin').trim() } };
    }
  }
  return { ok: false, error: 'invalid credentials' };
}

//  ITEMS 
function getItems() {
  return { ok: true, items: sheetRows('Items').filter(function(r){ return r.id && r.name; }) };
}

function saveItems(items, clearOnly) {
  var sh = getSheet('Items');
  var last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
  if (!clearOnly && Array.isArray(items) && items.length > 0) {
    var rows = items.map(function(d) {
      return [d.id||'', d.name||'', d.code||'', d.price||'', d.brand||'', d.folder||'all'];
    });
    sh.getRange(2, 1, rows.length, 6).setValues(rows);
  }
  return { ok: true, saved: clearOnly ? 0 : (items ? items.length : 0) };
}

function appendItems(items) {
  if (!Array.isArray(items) || !items.length) return { ok: true, appended: 0 };
  var sh = getSheet('Items');
  var rows = items.map(function(d) {
    return [d.id||'', d.name||'', d.code||'', d.price||'', d.brand||'', d.folder||'all'];
  });
  // Append row by row (batch)
  var lastRow = sh.getLastRow();
  sh.getRange(lastRow + 1, 1, rows.length, 6).setValues(rows);
  return { ok: true, appended: rows.length };
}

function deleteItemById(id) {
  var sh = getSheet('Items');
  var vals = sh.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][0]) === String(id)) { sh.deleteRow(i + 1); break; }
  }
  return { ok: true };
}

//  FOLDERS 
function getFolders() {
  var sh = getSheet('Folders');
  var list = sh.getDataRange().getValues().slice(1)
    .map(function(r){ return String(r[0]); }).filter(Boolean);
  if (list.indexOf('all') < 0) list.unshift('all');
  return { ok: true, folders: list };
}

function saveFolders(folders) {
  var sh = getSheet('Folders');
  var last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
  var rows = (folders||[]).filter(function(f){ return f !== 'all'; })
    .map(function(f){ return [f]; });
  if (rows.length > 0) sh.getRange(2, 1, rows.length, 1).setValues(rows);
  return { ok: true };
}

//  QUOTATIONS 
function getHistory() {
  return { ok: true, history: sheetRows('Quotations').filter(function(r){ return r.id; })
    .map(function(r) {
      return { id:r.id, date:r.date, client:r.client, qno:r.qno, site:r.site,
               rows: safeJSON(r.rows_json, []),
               totalBefore:r.total_before, totalTax:r.total_tax, totalAfter:r.total_after };
    })};
}

function saveQuotation(q) {
  if (!q || !q.id) return { ok: false, error: 'no data' };
  var sh = getSheet('Quotations');
  var row = [q.id, q.date||'', q.client||'', q.qno||'', q.site||'',
             JSON.stringify(q.rows||[]), q.totalBefore||'', q.totalTax||'', q.totalAfter||''];
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(q.id)) {
      sh.getRange(i+1, 1, 1, 9).setValues([row]);
      return { ok: true, action: 'updated' };
    }
  }
  sh.appendRow(row);
  return { ok: true, action: 'created' };
}

function deleteQuotation(id) {
  var sh = getSheet('Quotations');
  var vals = sh.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][0]) === String(id)) { sh.deleteRow(i + 1); break; }
  }
  return { ok: true };
}

function clearHistory() {
  var sh = getSheet('Quotations');
  var last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
  return { ok: true };
}

//  CONFIG 
function getConfig() {
  var rows = sheetRows('Config');
  var cfg = {};
  rows.forEach(function(r){ if (r.key) cfg[String(r.key)] = r.value; });
  return { ok: true, config: cfg };
}

function setConfig(key, value) {
  if (!key) return { ok: false, error: 'no key' };
  var sh = getSheet('Config');
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] === key) { sh.getRange(i+1, 2).setValue(value); return { ok: true }; }
  }
  sh.appendRow([key, value]);
  return { ok: true };
}

//  UTIL 
function safeJSON(str, fallback) {
  if (!str) return fallback;
  if (typeof str === 'object') return str;
  try { return JSON.parse(str); } catch(e) { return fallback; }
}

//  DEVICE TOKEN MANAGEMENT 
function getKey(token) {
  if (!token) return { ok: false, error: 'no token' };
  
  var sh = getSheet('Devices');
  var vals = sh.getDataRange().getValues();
  
  // Find this token in approved devices
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(token).trim() && vals[i][2] === true) {
      // Approved - return the encrypted key
      var cfg = getConfig();
      if (cfg.ok && cfg.config.gem_enc) {
        return { ok: true, key: cfg.config.gem_enc };
      }
      return { ok: false, error: 'key not configured' };
    }
    if (String(vals[i][0]).trim() === String(token).trim() && vals[i][2] !== true) {
      return { ok: false, pending: true };
    }
  }
  // Token not found - register as pending
  sh.appendRow([token, '', false, new Date().toISOString()]);
  return { ok: false, pending: true };
}

function registerDevice(token, label) {
  if (!token) return { ok: false };
  var sh  = getSheet('Devices');
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(token).trim()) {
      sh.getRange(i+1, 2).setValue(label || '');
      sh.getRange(i+1, 5).setValue(new Date().toISOString());
      return { ok: true, updated: true };
    }
  }
  sh.appendRow([token, label || '', false, new Date().toISOString(), new Date().toISOString()]);
  return { ok: true, created: true };
}
