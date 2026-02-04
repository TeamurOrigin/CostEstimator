var ESTIMATES_SHEET_NAME = 'Сметы';
var REF_SHEET_NAME = 'Справочник сметы';
var ITEMS_SHEET_PREFIX = 'Смета';
var ESTIMATES_STORE_KEY = 'estimates_store_v1';
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
    .setTitle('Конструктор сметы')
    .setWidth(760)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Конструктор сметы');
}

function openEstimateItems(id) {
  ensureCoreSheets_();

  var boot = getEstimateItemsBootstrap(id);

  var template = HtmlService.createTemplateFromFile('EstimateItemsModal');
  template.bootstrap = boot;
  template.estimateId = String(id);

  var html = template.evaluate()
    .setTitle('Позиции сметы')
    .setWidth(1100)
    .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(html, 'Позиции сметы');
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
  var category = String(payload.category || '').trim();
  var type = String(payload.type || '').trim();
  if (!name) throw new Error('Название обязательно.');
  if (!category) throw new Error('Категория обязательна.');
  if (!type) throw new Error('Тип обязателен.');

  var id = makeId_();
  var now = new Date();

  var est = {
    id: id,
    name: name,
    category: category,
    type: type,
    createdAt: now.toISOString(),
    itemsCount: 0,
    totalSum: 0
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

  var totalSum = 0;
  for (var i = 0; i < items.length; i++) {
    var s = Number(items[i].rowSum || 0);
    if (isFinite(s)) totalSum += s;
  }

  est.itemsCount = items.length;
  est.totalSum = totalSum;

  return {
    estimate: est,
    items: items,
    articlesByCategory: ref.articlesByCategory
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
    if (!row.article && row.qty === 0 && row.unitCost === 0) continue;
    row.rowSum = calcRowSum_(row);
    norm.push(row);
  }

  writeItems_(estimateId, norm);

  var totalSum = 0;
  for (var j = 0; j < norm.length; j++) totalSum += Number(norm[j].rowSum || 0);

  est.itemsCount = norm.length;
  est.totalSum = totalSum;
  upsertEstimate_(est);

  return { itemsCount: est.itemsCount, totalSum: est.totalSum };
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

  var headers = ['Статья','Кол-во','Залов','Дней','Дней мероприятий','Коэф.','Стоимость за ед.','Сумма'];
  sheet.getRange(1,1,1,headers.length).setValues([headers]);

  var out = [];
  for (var i = 0; i < items.length; i++) {
    var r = items[i];
    out.push([
      r.article || '',
      r.qty || 0,
      r.halls || 0,
      r.days || 0,
      r.eventDays || 0,
      r.coef || 1,
      r.unitCost || 0,
      r.rowSum || 0
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
    estSheet.getRange(1,1,1,6).setValues([['ID','Название','Категория','Тип','Статей','Сумма']]);
  }

  var refSheet = ss.getSheetByName(REF_SHEET_NAME);
  if (!refSheet) {
    refSheet = ss.insertSheet(REF_SHEET_NAME);
    refSheet.getRange(1,1,1,4).setValues([['Категория','Тип','Статья','Цена']]);
  }
}

function getRefData_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) {
    return { categories: [], types: [], articlesByCategory: {} };
  }

  var vals = sh.getRange(2,1,last-1,4).getValues();

  var cats = new Set();
  var types = new Set();
  var map = {};

  for (var i = 0; i < vals.length; i++) {
    var c = String(vals[i][0] || '').trim();
    var t = String(vals[i][1] || '').trim();
    var a = String(vals[i][2] || '').trim();
    if (c) cats.add(c);
    if (t) types.add(t);
    if (c && a) {
      if (!map[c]) map[c] = [];
      map[c].push(a);
    }
  }

  var categories = Array.from(cats);
  var typeList = Array.from(types);

  for (var k in map) {
    map[k] = uniq_(map[k]);
  }

  return { categories: categories, types: typeList, articlesByCategory: map };
}

function listEstimates_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(ESTIMATES_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2,1,last-1,6).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var id = String(vals[i][0] || '').trim();
    if (!id) continue;
    out.push({
      id: id,
      name: vals[i][1] || '',
      category: vals[i][2] || '',
      type: vals[i][3] || '',
      itemsCount: Number(vals[i][4] || 0),
      totalSum: Number(vals[i][5] || 0)
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
  var vals = sh.getRange(2,1,last-1,6).getValues();
  for (var i = 0; i < vals.length; i++) {
    var rid = String(vals[i][0] || '').trim();
    if (rid === id) {
      return {
        id: rid,
        name: vals[i][1] || '',
        category: vals[i][2] || '',
        type: vals[i][3] || '',
        itemsCount: Number(vals[i][4] || 0),
        totalSum: Number(vals[i][5] || 0)
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
    sh.appendRow([est.id, est.name, est.category, est.type, est.itemsCount || 0, est.totalSum || 0]);
    return;
  }

  var vals = sh.getRange(2,1,last-1,1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var rid = String(vals[i][0] || '').trim();
    if (rid === est.id) {
      sh.getRange(i+2,1,1,6).setValues([[
        est.id,
        est.name,
        est.category,
        est.type,
        Number(est.itemsCount || 0),
        Number(est.totalSum || 0)
      ]]);
      return;
    }
  }

  sh.appendRow([est.id, est.name, est.category, est.type, Number(est.itemsCount||0), Number(est.totalSum||0)]);
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
    sh.getRange(1,1,1,8).setValues([['Статья','Кол-во','Залов','Дней','Дней мероприятий','Коэф.','Стоимость за ед.','Сумма']]);
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
    article: String(it.article || '').trim(),
    qty: toNum_(it.qty, 0),
    halls: toNum_(it.halls, 0),
    days: toNum_(it.days, 0),
    eventDays: toNum_(it.eventDays, 0),
    coef: toNum_(it.coef, 1),
    unitCost: toNum_(it.unitCost, 0),
    rowSum: 0
  };
}

function calcRowSum_(row) {
  return toNum_(row.qty,0) * toNum_(row.halls,0) * toNum_(row.days,0) * toNum_(row.eventDays,0) * toNum_(row.coef,1) * toNum_(row.unitCost,0);
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
