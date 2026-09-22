/**
 * AppADay 138: Ferment Log backend.
 * Bind this script to a Google Sheet (Extensions > Apps Script), run setup(),
 * then deploy as a Web app executing as Me with access Anyone.
 * Every request must carry the TOKEN stored in Script Properties.
 */

var TABS = {
  Inventory: ['ID', 'Name', 'Type', 'Unit', 'OnHand', 'ReorderAt', 'UpdatedAt', 'Par'],
  Batches: ['ID', 'Number', 'Name', 'Style', 'VolumeGal', 'TargetOG', 'TargetFG', 'Stage', 'TankID', 'CreatedAt'],
  BatchIngredients: ['ID', 'BatchID', 'InventoryID', 'Qty', 'Unit'],
  Readings: ['ID', 'BatchID', 'Date', 'Gravity', 'TempF', 'Note'],
  StageLog: ['ID', 'BatchID', 'Stage', 'Date'],
  Tanks: ['ID', 'Name', 'CapacityGal', 'BatchID', 'AssignedDate', 'Status'],
  Packages: ['ID', 'BatchID', 'Format', 'UnitLabel', 'UnitVolGal', 'UnitCount', 'OnHandUnits', 'Date']
};

var STAGES = ['Planned', 'Brewed', 'Fermenting', 'Crashing', 'Carbonating', 'Packaged', 'Kicked'];
var TYPES = ['Grain', 'Hop', 'Yeast', 'Adjunct'];
var FORMATS = ['Keg', 'Bottle', 'Can'];
var TANKABLE = ['Planned', 'Brewed', 'Fermenting', 'Crashing', 'Carbonating'];
var READABLE = ['Brewed', 'Fermenting', 'Crashing', 'Carbonating'];
var TANK_STATUSES = ['Dirty', 'Clean', 'Sanitized'];

var ID_PREFIX = {
  Inventory: 'INV',
  Batches: 'BAT',
  BatchIngredients: 'BI',
  Readings: 'RD',
  StageLog: 'SL',
  Tanks: 'TK',
  Packages: 'PK'
};

var DATE_COLS = ['UpdatedAt', 'CreatedAt', 'Date', 'AssignedDate'];
var NUM_COLS = ['OnHand', 'ReorderAt', 'Par', 'Number', 'VolumeGal', 'TargetOG', 'TargetFG', 'Qty',
  'Gravity', 'TempF', 'CapacityGal', 'UnitVolGal', 'UnitCount', 'OnHandUnits'];

/* ---------- Setup ---------- */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var headers = TABS[name];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    var rows = Math.max(sh.getMaxRows() - 1, 1);
    headers.forEach(function (h, i) {
      if (NUM_COLS.indexOf(h) < 0) sh.getRange(2, i + 1, rows, 1).setNumberFormat('@');
    });
  });
  migrate_();
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('TOKEN', token);
  }
  Logger.log('TOKEN: ' + token);
  return token;
}

/* Upgrades sheets built before the Crashing and Carbonating split and tank status. */
function migrate_() {
  var n = 0;
  readTable_('Batches').forEach(function (b) {
    if (b.Stage === 'Conditioning') { updateObj_('Batches', b.ID, { Stage: 'Crashing' }); n++; }
  });
  readTable_('StageLog').forEach(function (l) {
    if (l.Stage === 'Conditioning') { updateObj_('StageLog', l.ID, { Stage: 'Crashing' }); n++; }
  });
  readTable_('Tanks').forEach(function (t) {
    if (TANK_STATUSES.indexOf(t.Status) < 0) { updateObj_('Tanks', t.ID, { Status: 'Clean' }); n++; }
  });
  if (n) Logger.log('Migrated ' + n + ' rows.');
}

/* ---------- Helpers ---------- */

function sheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Missing tab ' + name + '. Run setup() in the script editor.');
  return sh;
}

function cellOut_(h, v) {
  if (NUM_COLS.indexOf(h) >= 0) {
    if (h === 'TempF' && (v === '' || v === null)) return '';
    return Number(v) || 0;
  }
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

function readTable_(name) {
  var sh = sheet_(name);
  var headers = TABS[name];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, headers.length).getValues();
  var out = [];
  values.forEach(function (row) {
    if (row[0] === '' || row[0] === null) return;
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = cellOut_(h, row[i]); });
    out.push(obj);
  });
  return out;
}

function findRow_(name, id) {
  var sh = sheet_(name);
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function setTextFormats_(sh, name, row) {
  TABS[name].forEach(function (h, i) {
    if (NUM_COLS.indexOf(h) < 0) sh.getRange(row, i + 1).setNumberFormat('@');
  });
}

function appendObj_(name, obj) {
  var sh = sheet_(name);
  var headers = TABS[name];
  var row = sh.getLastRow() + 1;
  setTextFormats_(sh, name, row);
  var values = headers.map(function (h) { return obj[h] === undefined || obj[h] === null ? '' : obj[h]; });
  sh.getRange(row, 1, 1, headers.length).setValues([values]);
  return obj;
}

function updateObj_(name, id, fields) {
  var sh = sheet_(name);
  var headers = TABS[name];
  var row = findRow_(name, id);
  if (row < 0) throw new Error(name + ' record ' + id + ' was not found.');
  var range = sh.getRange(row, 1, 1, headers.length);
  var values = range.getValues()[0];
  headers.forEach(function (h, i) {
    if (Object.prototype.hasOwnProperty.call(fields, h) && h !== 'ID') values[i] = fields[h];
  });
  setTextFormats_(sh, name, row);
  range.setValues([values]);
}

function deleteObj_(name, id) {
  var row = findRow_(name, id);
  if (row < 0) throw new Error(name + ' record ' + id + ' was not found.');
  sheet_(name).deleteRow(row);
}

function newId_(prefix) {
  return prefix + '-' + Utilities.getUuid().slice(0, 8);
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function fullState_() {
  var state = {};
  Object.keys(TABS).forEach(function (name) { state[name] = readTable_(name); });
  return state;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function round3_(n) {
  return Math.round(n * 1000) / 1000;
}

function checkToken_(token) {
  var stored = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!stored) throw new Error('No token found. Run setup() in the script editor first.');
  return String(token || '') === stored;
}

/* ---------- Validation ---------- */

function fail_(msg) { throw new Error(msg); }

function reqStr_(d, k) {
  var v = d[k] === undefined || d[k] === null ? '' : String(d[k]).trim();
  if (!v) fail_(k + ' is required.');
  return v;
}

function optStr_(d, k) {
  return d[k] === undefined || d[k] === null ? '' : String(d[k]).trim();
}

function reqNum_(d, k, min, max) {
  var raw = d[k];
  var v = Number(raw);
  if (raw === '' || raw === null || raw === undefined || !isFinite(v)) fail_(k + ' must be a number.');
  if (min !== undefined && min !== null && v < min) fail_(k + ' must be at least ' + min + '.');
  if (max !== undefined && max !== null && v > max) fail_(k + ' must be ' + max + ' or less.');
  return v;
}

function reqInt_(d, k, min) {
  var v = reqNum_(d, k, min);
  if (Math.floor(v) !== v) fail_(k + ' must be a whole number.');
  return v;
}

function reqEnum_(d, k, list) {
  var v = reqStr_(d, k);
  if (list.indexOf(v) < 0) fail_(k + ' must be one of ' + list.join(', ') + '.');
  return v;
}

function date_(d, k) {
  var v = optStr_(d, k);
  if (!v) return today_();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) fail_(k + ' must be a date in YYYY-MM-DD form.');
  return v;
}

function par_(d, reorderAt) {
  var raw = d.Par;
  if (raw === '' || raw === null || raw === undefined) return 0;
  var v = round3_(reqNum_(d, 'Par', 0));
  if (v > 0 && v < reorderAt) fail_('Par must be at or above the reorder level.');
  return v;
}

function getById_(rows, id, label) {
  for (var i = 0; i < rows.length; i++) if (rows[i].ID === String(id)) return rows[i];
  fail_(label + ' ' + id + ' was not found.');
}

/* ---------- Actions ---------- */

var ACTIONS = {
  addInventory: function (d) {
    var rec = {
      ID: newId_(ID_PREFIX.Inventory),
      Name: reqStr_(d, 'Name'),
      Type: reqEnum_(d, 'Type', TYPES),
      Unit: reqStr_(d, 'Unit'),
      OnHand: round3_(reqNum_(d, 'OnHand', 0)),
      ReorderAt: round3_(reqNum_(d, 'ReorderAt', 0)),
      UpdatedAt: today_()
    };
    rec.Par = par_(d, rec.ReorderAt);
    appendObj_('Inventory', rec);
  },

  updateInventory: function (d) {
    var id = reqStr_(d, 'ID');
    getById_(readTable_('Inventory'), id, 'Inventory item');
    var fields = {
      Name: reqStr_(d, 'Name'),
      Type: reqEnum_(d, 'Type', TYPES),
      Unit: reqStr_(d, 'Unit'),
      ReorderAt: round3_(reqNum_(d, 'ReorderAt', 0)),
      UpdatedAt: today_()
    };
    fields.Par = par_(d, fields.ReorderAt);
    updateObj_('Inventory', id, fields);
  },

  adjustInventory: function (d) {
    var id = reqStr_(d, 'ID');
    var item = getById_(readTable_('Inventory'), id, 'Inventory item');
    var delta = reqNum_(d, 'Delta');
    var next = round3_(item.OnHand + delta);
    if (next < 0) fail_(item.Name + ' cannot go below 0 ' + item.Unit + '.');
    updateObj_('Inventory', id, { OnHand: next, UpdatedAt: today_() });
  },

  createBatch: function (d) {
    var number = reqInt_(d, 'Number', 1);
    var name = reqStr_(d, 'Name');
    var style = reqStr_(d, 'Style');
    var vol = reqNum_(d, 'VolumeGal', 0.1);
    var og = reqNum_(d, 'TargetOG', 1.0, 1.2);
    var fg = reqNum_(d, 'TargetFG', 0.98, 1.2);
    if (fg >= og) fail_('TargetFG must be lower than TargetOG.');
    var batches = readTable_('Batches');
    batches.forEach(function (b) { if (b.Number === number) fail_('Batch number ' + number + ' is already used.'); });

    var inv = readTable_('Inventory');
    var ings = Array.isArray(d.Ingredients) ? d.Ingredients : [];
    var totals = {};
    var lines = ings.map(function (ing) {
      var invId = reqStr_(ing, 'InventoryID');
      var qty = reqNum_(ing, 'Qty', 0);
      if (qty <= 0) fail_('Every ingredient quantity must be above 0.');
      var item = getById_(inv, invId, 'Inventory item');
      totals[invId] = round3_((totals[invId] || 0) + qty);
      return { item: item, qty: qty };
    });
    Object.keys(totals).forEach(function (invId) {
      var item = getById_(inv, invId, 'Inventory item');
      if (totals[invId] > item.OnHand + 1e-9) {
        fail_('Not enough ' + item.Name + '. Needs ' + totals[invId] + ' ' + item.Unit + ', has ' + item.OnHand + '.');
      }
    });

    var batchId = newId_(ID_PREFIX.Batches);
    var today = today_();
    appendObj_('Batches', {
      ID: batchId, Number: number, Name: name, Style: style, VolumeGal: vol,
      TargetOG: og, TargetFG: fg, Stage: 'Planned', TankID: '', CreatedAt: today
    });
    lines.forEach(function (l) {
      appendObj_('BatchIngredients', {
        ID: newId_(ID_PREFIX.BatchIngredients), BatchID: batchId,
        InventoryID: l.item.ID, Qty: l.qty, Unit: l.item.Unit
      });
    });
    Object.keys(totals).forEach(function (invId) {
      var item = getById_(inv, invId, 'Inventory item');
      updateObj_('Inventory', invId, { OnHand: round3_(item.OnHand - totals[invId]), UpdatedAt: today });
    });
    appendObj_('StageLog', { ID: newId_(ID_PREFIX.StageLog), BatchID: batchId, Stage: 'Planned', Date: today });
  },

  logReading: function (d) {
    var batch = getById_(readTable_('Batches'), reqStr_(d, 'BatchID'), 'Batch');
    if (READABLE.indexOf(batch.Stage) < 0) fail_('Readings can only be logged for Brewed, Fermenting, Crashing, or Carbonating batches.');
    var temp = optStr_(d, 'TempF');
    var rec = {
      ID: newId_(ID_PREFIX.Readings),
      BatchID: batch.ID,
      Date: date_(d, 'Date'),
      Gravity: reqNum_(d, 'Gravity', 0.98, 1.2),
      TempF: temp === '' ? '' : reqNum_(d, 'TempF', 20, 220),
      Note: optStr_(d, 'Note')
    };
    appendObj_('Readings', rec);
  },

  deleteReading: function (d) {
    var id = reqStr_(d, 'ID');
    getById_(readTable_('Readings'), id, 'Reading');
    deleteObj_('Readings', id);
  },

  advanceStage: function (d) {
    var batch = getById_(readTable_('Batches'), reqStr_(d, 'BatchID'), 'Batch');
    var date = date_(d, 'Date');
    var idx = STAGES.indexOf(batch.Stage);
    if (idx < 0) fail_('Batch has an unknown stage: ' + batch.Stage + '.');
    if (batch.Stage === 'Kicked') fail_('This batch is already kicked.');
    var next = STAGES[idx + 1];
    if (next === 'Packaged') fail_('Package this batch from the Packaged tab to move it to Packaged.');
    updateObj_('Batches', batch.ID, { Stage: next });
    appendObj_('StageLog', { ID: newId_(ID_PREFIX.StageLog), BatchID: batch.ID, Stage: next, Date: date });
  },

  addTank: function (d) {
    appendObj_('Tanks', {
      ID: newId_(ID_PREFIX.Tanks),
      Name: reqStr_(d, 'Name'),
      CapacityGal: reqNum_(d, 'CapacityGal', 0.1),
      BatchID: '',
      AssignedDate: '',
      Status: optStr_(d, 'Status') ? reqEnum_(d, 'Status', TANK_STATUSES) : 'Clean'
    });
  },

  setTankStatus: function (d) {
    var tank = getById_(readTable_('Tanks'), reqStr_(d, 'TankID'), 'Tank');
    var status = reqEnum_(d, 'Status', TANK_STATUSES);
    if (tank.BatchID) fail_(tank.Name + ' holds a batch. Release it before changing its status.');
    updateObj_('Tanks', tank.ID, { Status: status });
  },

  assignTank: function (d) {
    var tank = getById_(readTable_('Tanks'), reqStr_(d, 'TankID'), 'Tank');
    var batch = getById_(readTable_('Batches'), reqStr_(d, 'BatchID'), 'Batch');
    var date = date_(d, 'Date');
    if (tank.BatchID) fail_(tank.Name + ' already holds a batch. Release it first.');
    if (batch.TankID) fail_('Batch ' + batch.Number + ' is already in a tank.');
    if (tank.Status !== 'Sanitized') fail_(tank.Name + ' is not sanitized. Mark it sanitized before filling it.');
    if (TANKABLE.indexOf(batch.Stage) < 0) fail_('Only Planned through Carbonating batches can go in a tank.');
    if (batch.VolumeGal > tank.CapacityGal) fail_('Batch ' + batch.Number + ' is larger than ' + tank.Name + '.');
    updateObj_('Tanks', tank.ID, { BatchID: batch.ID, AssignedDate: date });
    updateObj_('Batches', batch.ID, { TankID: tank.ID });
  },

  releaseTank: function (d) {
    var tank = getById_(readTable_('Tanks'), reqStr_(d, 'TankID'), 'Tank');
    if (!tank.BatchID) fail_(tank.Name + ' is already empty.');
    var batchRow = findRow_('Batches', tank.BatchID);
    updateObj_('Tanks', tank.ID, { BatchID: '', AssignedDate: '', Status: 'Dirty' });
    if (batchRow > 0) updateObj_('Batches', tank.BatchID, { TankID: '' });
  },

  packageBatch: function (d) {
    var batch = getById_(readTable_('Batches'), reqStr_(d, 'BatchID'), 'Batch');
    var date = date_(d, 'Date');
    if (batch.Stage !== 'Carbonating') fail_('Only Carbonating batches can be packaged.');
    var list = Array.isArray(d.Packages) ? d.Packages : [];
    if (!list.length) fail_('Add at least one package format.');
    var recs = list.map(function (p) {
      var count = reqInt_(p, 'UnitCount', 1);
      return {
        ID: newId_(ID_PREFIX.Packages),
        BatchID: batch.ID,
        Format: reqEnum_(p, 'Format', FORMATS),
        UnitLabel: reqStr_(p, 'UnitLabel'),
        UnitVolGal: reqNum_(p, 'UnitVolGal', 0.001),
        UnitCount: count,
        OnHandUnits: count,
        Date: date
      };
    });
    var tanks = readTable_('Tanks');
    recs.forEach(function (r) { appendObj_('Packages', r); });
    updateObj_('Batches', batch.ID, { Stage: 'Packaged', TankID: '' });
    appendObj_('StageLog', { ID: newId_(ID_PREFIX.StageLog), BatchID: batch.ID, Stage: 'Packaged', Date: date });
    tanks.forEach(function (t) {
      if (t.BatchID === batch.ID) updateObj_('Tanks', t.ID, { BatchID: '', AssignedDate: '', Status: 'Dirty' });
    });
  },

  depleteGoods: function (d) {
    var pk = getById_(readTable_('Packages'), reqStr_(d, 'ID'), 'Package');
    var units = reqInt_(d, 'Units', 1);
    if (pk.OnHandUnits <= 0) fail_('No units of ' + pk.UnitLabel + ' are left.');
    var next = Math.max(0, pk.OnHandUnits - units);
    updateObj_('Packages', pk.ID, { OnHandUnits: next });
    var siblings = readTable_('Packages').filter(function (p) { return p.BatchID === pk.BatchID; });
    var remaining = siblings.reduce(function (s, p) { return s + p.OnHandUnits; }, 0);
    if (remaining <= 0) {
      var batchRow = findRow_('Batches', pk.BatchID);
      if (batchRow > 0) {
        var batch = getById_(readTable_('Batches'), pk.BatchID, 'Batch');
        if (batch.Stage === 'Packaged') {
          updateObj_('Batches', batch.ID, { Stage: 'Kicked' });
          appendObj_('StageLog', { ID: newId_(ID_PREFIX.StageLog), BatchID: batch.ID, Stage: 'Kicked', Date: today_() });
        }
      }
    }
  }
};

/* ---------- Web app entry points ---------- */

function doGet(e) {
  try {
    var token = e && e.parameter ? e.parameter.token : '';
    if (!checkToken_(token)) return json_({ ok: false, error: 'Invalid token. Check Settings.' });
    return json_({ ok: true, state: fullState_(), serverTime: new Date().toISOString() });
  } catch (err) {
    return json_({ ok: false, error: err && err.message ? err.message : String(err) });
  }
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Request body is not valid JSON.' });
  }
  var lock = LockService.getScriptLock();
  try {
    if (!checkToken_(body.token)) return json_({ ok: false, error: 'Invalid token. Check Settings.' });
    var handler = ACTIONS[body.action];
    if (!handler) return json_({ ok: false, error: 'Unknown action: ' + body.action });
    lock.waitLock(10000);
    handler(body.data || {});
    SpreadsheetApp.flush();
    return json_({ ok: true, state: fullState_(), serverTime: new Date().toISOString() });
  } catch (err) {
    return json_({ ok: false, error: err && err.message ? err.message : String(err) });
  } finally {
    lock.releaseLock();
  }
}
