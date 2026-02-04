var ESTIMATES_SHEET_NAME = 'Сметы';
var REF_SHEET_NAME = 'Справочник сметы';
var ITEMS_SHEET_PREFIX = 'Смета';
var ITEMS_STORE_PREFIX = 'estimate_items_v1_';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Конструктор смет')
    .addItem('Открыть', 'openEstimateBuilder')
    .addToUi();
}

function openEstimateBuilder() {
  ensureCoreSheets_();
  var html = HtmlService.createHtmlOutputFromFile('EstimateModal')
    .setTitle('Проекты')
    .setWidth(760)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Проекты');
}

function openEstimateItems(id) {
  ensureCoreSheets_();

  var boot = getEstimateItemsBootstrap(id);

  var template = HtmlService.createTemplateFromFile('EstimateItemsModal');
  template.bootstrap = boot;
  template.estimateId = String(id);

  var html = template.evaluate()
    .setTitle('Позиции проекта')
    .setWidth(980)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, 'Позиции проекта');
  return true;
}

function getBootstrapData() {
  ensureCoreSheets_();
  var estimates = listEstimates_();
  var ref = getRefData_();

  return {
    categories: ref.categories,
    types: ref.types,
    estimates: estimates
  };
}

function createEstimate(payload) {
  ensureCoreSheets_();
  if (!payload) payload = {};
  var name = String(payload.name || '').trim();
  var type = String(payload.type || '').trim();
  var group = String(payload.group || '').trim();
  if (!name) throw new Error('Название обязательно.');
  if (!type) throw new Error('Тип обязателен.');
  if (!group) throw new Error('Группа обязательна.');

  var id = makeId_();
  var now = new Date();

  var est = {
    id: id,
    name: name,
    type: type,
    group: group,
    createdAt: now.toISOString(),
    itemsCount: 0
  };

  upsertEstimate_(est);
  ensureItemsSheet_(id, name);

  return est;
}

function deleteEstimate(id) {
  ensureCoreSheets_();
  id = String(id || '').trim();
  if (!id) throw new Error('Некорректный ID.');

  deleteEstimate_(id);
  deleteItems_(id);

  var sheet = SpreadsheetApp.getActive().getSheetByName(itemsSheetName_(id));
  if (sheet) SpreadsheetApp.getActive().deleteSheet(sheet);

  return true;
}

function getEstimateItemsBootstrap(id) {
  ensureCoreSheets_();
  id = String(id || '').trim();
  if (!id) throw new Error('Смета не найдена.');

  var est = findEstimateById_(id);
  if (!est) throw new Error('Смета не найдена.');

  var ref = getRefData_();
  var items = readItems_(id);

  est.itemsCount = items.length;

  return {
    estimate: est,
    items: items,
    categories: ref.categories,
    positionsByCategory: ref.positionsByCategory,
    positions: ref.positions
  };
}

function saveEstimateItems(estimateId, items) {
  ensureCoreSheets_();
  estimateId = String(estimateId || '').trim();
  if (!estimateId) throw new Error('Смета не найдена.');

  var est = findEstimateById_(estimateId);
  if (!est) throw new Error('Смета не найдена.');

  if (!items || !items.length) items = [];

  var norm = [];
  for (var i = 0; i < items.length; i++) {
    var it = items[i] || {};
    var row = normalizeItemRow_(it);
    if (!row.category && !row.position && !row.comment) continue;
    norm.push(row);
  }

  writeItems_(estimateId, norm);

  est.itemsCount = norm.length;
  upsertEstimate_(est);

  return { itemsCount: est.itemsCount };
}

function exportEstimateToSheet(estimateId) {
  ensureCoreSheets_();
  estimateId = String(estimateId || '').trim();
  if (!estimateId) throw new Error('Смета не найдена.');

  var est = findEstimateById_(estimateId);
  if (!est) throw new Error('Смета не найдена.');

  var items = readItems_(estimateId);
  var sheet = ensureItemsSheet_(estimateId, est.name);

  sheet.clear();

  var headers = ['Категория','Позиция','Комментарий'];
  sheet.getRange(1,1,1,headers.length).setValues([headers]);

  var out = [];
  for (var i = 0; i < items.length; i++) {
    var r = items[i];
    out.push([
      r.category || '',
      r.position || '',
      r.comment || ''
    ]);
  }
  if (out.length) sheet.getRange(2,1,out.length,headers.length).setValues(out);

  sheet.autoResizeColumns(1, headers.length);

  return true;
}

/* =========================
   Storage + Sheets helpers
   ========================= */

function ensureCoreSheets_() {
  var ss = SpreadsheetApp.getActive();

  var estSheet = ss.getSheetByName(ESTIMATES_SHEET_NAME);
  if (!estSheet) {
    estSheet = ss.insertSheet(ESTIMATES_SHEET_NAME);
    estSheet.getRange(1,1,1,5).setValues([['ID','Название','Тип','Группа','Позиции']]);
  }

  var refSheet = ss.getSheetByName(REF_SHEET_NAME);
  if (!refSheet) {
    refSheet = ss.insertSheet(REF_SHEET_NAME);
    refSheet.getRange(1,1,1,4).setValues([['Категория','Позиция','Цена','Тип']]);
  }
}

function getRefData_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) {
    return { categories: [], types: [], positionsByCategory: {}, positions: [] };
  }

  var vals = sh.getRange(2,1,last-1,4).getValues();

  var cats = new Set();
  var types = new Set();
  var map = {};
  var positions = new Set();

  for (var i = 0; i < vals.length; i++) {
    var c = String(vals[i][0] || '').trim();
    var p = String(vals[i][1] || '').trim();
    var t = String(vals[i][3] || '').trim();
    if (c) cats.add(c);
    if (t) types.add(t);
    if (p) positions.add(p);
    if (c && p) {
      if (!map[c]) map[c] = [];
      map[c].push(p);
    }
  }

  var categories = Array.from(cats);
  var typeList = Array.from(types);
  var positionsList = Array.from(positions);

  for (var k in map) {
    map[k] = uniq_(map[k]);
  }

  return { categories: categories, types: typeList, positionsByCategory: map, positions: positionsList };
}

function listEstimates_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(ESTIMATES_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2,1,last-1,5).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var id = String(vals[i][0] || '').trim();
    if (!id) continue;
    out.push({
      id: id,
      name: vals[i][1] || '',
      type: vals[i][2] || '',
      group: vals[i][3] || '',
      itemsCount: Number(vals[i][4] || 0)
    });
  }
  return out;
}

function findEstimateById_(id) {
  id = String(id || '').trim();
  if (!id) return null;
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(ESTIMATES_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return null;
  var vals = sh.getRange(2,1,last-1,5).getValues();
  for (var i = 0; i < vals.length; i++) {
    var rid = String(vals[i][0] || '').trim();
    if (rid === id) {
      return {
        id: rid,
        name: vals[i][1] || '',
        type: vals[i][2] || '',
        group: vals[i][3] || '',
        itemsCount: Number(vals[i][4] || 0)
      };
    }
  }
  return null;
}

function upsertEstimate_(est) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(ESTIMATES_SHEET_NAME);

  var last = sh.getLastRow();
  if (last < 2) {
    sh.appendRow([est.id, est.name, est.type, est.group || '', est.itemsCount || 0]);
    return;
  }

  var vals = sh.getRange(2,1,last-1,1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var rid = String(vals[i][0] || '').trim();
    if (rid === est.id) {
      sh.getRange(i+2,1,1,5).setValues([[
        est.id,
        est.name,
        est.type,
        est.group || '',
        Number(est.itemsCount || 0)
      ]]);
      return;
    }
  }

  sh.appendRow([est.id, est.name, est.type, est.group || '', Number(est.itemsCount||0)]);
}

function deleteEstimate_(id) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(ESTIMATES_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return;

  var vals = sh.getRange(2,1,last-1,1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    var rid = String(vals[i][0] || '').trim();
    if (rid === id) {
      sh.deleteRow(i + 2);
      break;
    }
  }
}

function itemsSheetName_(id) {
  return ITEMS_SHEET_PREFIX + ' ' + String(id).slice(0, 8);
}

function ensureItemsSheet_(id, estName) {
  var ss = SpreadsheetApp.getActive();
  var name = itemsSheetName_(id);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,3).setValues([['Категория','Позиция','Комментарий']]);
    sh.getRange(1,10).setValue('ID');
    sh.getRange(1,11).setValue(String(id));
    if (estName) sh.getRange(1,12).setValue(String(estName));
  }
  return sh;
}

/* =========================
   Items storage
   ========================= */

function readItems_(estimateId) {
  var props = PropertiesService.getDocumentProperties();
  var raw = props.getProperty(ITEMS_STORE_PREFIX + estimateId);
  if (!raw) return [];
  var data = JSON.parse(raw);
  if (!data || !data.length) return [];
  return data;
}

function writeItems_(estimateId, items) {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty(ITEMS_STORE_PREFIX + estimateId, JSON.stringify(items || []));
}

function deleteItems_(estimateId) {
  var props = PropertiesService.getDocumentProperties();
  props.deleteProperty(ITEMS_STORE_PREFIX + estimateId);
}

function normalizeItemRow_(it) {
  return {
    category: String(it.category || '').trim(),
    position: String(it.position || '').trim(),
    comment: String(it.comment || '').trim()
  };
}

/* =========================
   Utils
   ========================= */

function makeId_() {
  var s = Utilities.getUuid().replace(/-/g,'');
  return s;
}

function toNum_(v, def) {
  if (def === undefined) def = 0;
  if (v === null || v === undefined || v === '') return def;
  var x = Number(String(v).replace(',', '.'));
  return isFinite(x) ? x : def;
}

function uniq_(arr) {
  var seen = {};
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    var v = String(arr[i] || '').trim();
    if (!v) continue;
    if (seen[v]) continue;
    seen[v] = true;
    out.push(v);
  }
  return out;
}
