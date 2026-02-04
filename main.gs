var ESTIMATES_SHEET_NAME = 'Сметы';
var REF_SHEET_NAME = 'Справочник сметы';
var ITEMS_SHEET_PREFIX = 'Смета';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Смета')
    .addItem('Конструктор сметы', 'openEstimateBuilder')
    .addToUi();
}

function openEstimateBuilder() {
  ensureCoreSheets_();
  var html = HtmlService.createHtmlOutputFromFile('EstimateModal')
    .setTitle('Конструктор сметы')
    .setWidth(1150)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Конструктор сметы');
}

function getBootstrapData() {
  ensureCoreSheets_();
  var a = readRef_();
  return { categories: a.categories, articlesByCategory: a.articlesByCategory, estimates: listEstimates_() };
}

function createEstimate(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();

    var name = String(payload && payload.name ? payload.name : '').trim();
    var category = String(payload && payload.category ? payload.category : '').trim();
    if (!name) throw new Error('Укажите название.');
    if (!category) throw new Error('Укажите категорию.');

    var ss = SpreadsheetApp.getActive();
    var estSheet = ss.getSheetByName(ESTIMATES_SHEET_NAME);

    var data = estSheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][1]).trim() === name && String(data[r][2]).trim() === category) {
        return { id: String(data[r][0]), name: name, category: category, itemsCount: Number(data[r][3]) || 0, totalSum: Number(data[r][4]) || 0, sheetName: String(data[r][5] || '') };
      }
    }

    var id = Utilities.getUuid();
    var baseSheetName = ITEMS_SHEET_PREFIX + ' • ' + category + ' • ' + name;
    var sheetName = makeUniqueSheetName_(baseSheetName, ss);

    var itemsSheet = ss.insertSheet(sheetName);
    buildItemsSheet_(itemsSheet, name, category);

    estSheet.appendRow([id, name, category, 0, 0, sheetName, new Date()]);
    try { estSheet.hideColumns(1); } catch (e) {}

    return { id: id, name: name, category: category, itemsCount: 0, totalSum: 0, sheetName: sheetName };
  } finally {
    lock.releaseLock();
  }
}

function deleteEstimate(id) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();

    var ss = SpreadsheetApp.getActive();
    var estSheet = ss.getSheetByName(ESTIMATES_SHEET_NAME);
    var data = estSheet.getDataRange().getValues();

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][0]) === String(id)) {
        var sheetName = String(data[r][5] || '').trim();
        estSheet.deleteRow(r + 1);
        if (sheetName) {
          var sh = ss.getSheetByName(sheetName);
          if (sh) ss.deleteSheet(sh);
        }
        return true;
      }
    }
    return false;
  } finally {
    lock.releaseLock();
  }
}

function getEstimateItems(id) {
  ensureCoreSheets_();

  var est = findEstimateById_(id);
  if (!est) throw new Error('Смета не найдена.');

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(est.sheetName);
  if (!sh) throw new Error('Лист сметы не найден.');

  var last = sh.getLastRow();
  if (last < 3) return { estimate: est, items: [] };

  var values = sh.getRange(3, 1, last - 2, 7).getValues();
  var items = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var article = String(row[0] || '').trim();
    var qty = toNumber_(row[1], 0);
    var halls = toNumber_(row[2], 0);
    var days = toNumber_(row[3], 0);
    var eventDays = toNumber_(row[4], 0);
    var coef = toNumber_(row[5], 1);
    var unitCost = toNumber_(row[6], 0);

    if (!article && qty === 0 && halls === 0 && days === 0 && eventDays === 0 && unitCost === 0) continue;

    items.push({ article: article, qty: qty, halls: halls, days: days, eventDays: eventDays, coef: coef, unitCost: unitCost });
  }

  return { estimate: est, items: items };
}

function saveEstimateItems(id, items) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();

    var est = findEstimateById_(id);
    if (!est) throw new Error('Смета не найдена.');

    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(est.sheetName);
    if (!sh) throw new Error('Лист сметы не найден.');

    var cleaned = Array.isArray(items) ? items : [];
    var out = [];
    var total = 0;
    var count = 0;

    for (var i = 0; i < cleaned.length; i++) {
      var it = cleaned[i] || {};
      var article = String(it.article || '').trim();
      if (!article) continue;

      var qty = toNumber_(it.qty, 0);
      var halls = toNumber_(it.halls, 0);
      var days = toNumber_(it.days, 0);
      var eventDays = toNumber_(it.eventDays, 0);
      var coef = toNumber_(it.coef, 1);
      var unitCost = toNumber_(it.unitCost, 0);

      out.push([article, qty, halls, days, eventDays, coef, unitCost]);
      total += qty * halls * days * eventDays * coef * unitCost;
      count++;
    }

    var lastRow = Math.max(sh.getLastRow(), 3);
    if (lastRow >= 3) sh.getRange(3, 1, lastRow - 2, 8).clearContent();

    if (out.length) {
      sh.getRange(3, 1, out.length, 7).setValues(out);

      var formulas = [];
      for (var k = 0; k < out.length; k++) {
        var rr = 3 + k;
        formulas.push(['=B' + rr + '*C' + rr + '*D' + rr + '*E' + rr + '*F' + rr + '*G' + rr]);
      }
      sh.getRange(3, 8, formulas.length, 1).setFormulas(formulas);

      sh.getRange(3, 2, out.length, 5).setNumberFormat('0.########');
      sh.getRange(3, 7, out.length, 2).setNumberFormat('0.00');

      applyArticleValidation_(sh, est.category, out.length);
    } else {
      applyArticleValidation_(sh, est.category, 1);
    }

    updateEstimateTotals_(id, count, total, est.sheetName);
    return { itemsCount: count, totalSum: total };
  } finally {
    lock.releaseLock();
  }
}

function activateEstimateSheet(id) {
  ensureCoreSheets_();

  var est = findEstimateById_(id);
  if (!est) throw new Error('Смета не найдена.');

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(est.sheetName);
  if (!sh) throw new Error('Лист сметы не найден.');

  ss.setActiveSheet(sh);
  return true;
}

function ensureCoreSheets_() {
  var ss = SpreadsheetApp.getActive();

  var est = ss.getSheetByName(ESTIMATES_SHEET_NAME);
  if (!est) {
    est = ss.insertSheet(ESTIMATES_SHEET_NAME);
    est.getRange(1, 1, 1, 7).setValues([['ID', 'Название', 'Категория', 'Итого статей', 'Итого сумма', 'Лист', 'Обновлено']]);
    est.setFrozenRows(1);
    est.getRange(1, 2, 1, 6).setFontWeight('bold');
    try { est.hideColumns(1); } catch (e) {}
    est.autoResizeColumns(2, 6);
  }

  if (!ss.getSheetByName(REF_SHEET_NAME)) {
    var ref = ss.insertSheet(REF_SHEET_NAME);
    ref.getRange('A1:B1').setValues([['Категория', 'Статья']]);
    ref.setFrozenRows(1);
    ref.autoResizeColumns(1, 2);
  }
}

function readRef_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF_SHEET_NAME);
  if (!sh) return { categories: [], articlesByCategory: {} };

  var last = sh.getLastRow();
  if (last < 2) return { categories: [], articlesByCategory: {} };

  var cats = sh.getRange(2, 1, last - 1, 1).getValues();
  var arts = sh.getRange(2, 2, last - 1, 1).getValues();

  var map = {};
  for (var i = 0; i < cats.length; i++) {
    var c = String(cats[i][0] || '').trim();
    var a = String(arts[i][0] || '').trim();
    if (!c || !a) continue;
    if (!map[c]) map[c] = [];
    map[c].push(a);
  }

  var categories = Object.keys(map).sort(function(x, y) { return x.localeCompare(y, 'ru'); });

  for (var j = 0; j < categories.length; j++) {
    var key = categories[j];
    var seen = {};
    var uniq = [];
    for (var k = 0; k < map[key].length; k++) {
      var v = map[key][k];
      if (seen[v]) continue;
      seen[v] = true;
      uniq.push(v);
    }
    uniq.sort(function(x, y) { return x.localeCompare(y, 'ru'); });
    map[key] = uniq;
  }

  return { categories: categories, articlesByCategory: map };
}

function listEstimates_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(ESTIMATES_SHEET_NAME);
  if (!sh) return [];

  var values = sh.getDataRange().getValues();
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var id = String(row[0] || '').trim();
    if (!id) continue;
    out.push({ id: id, name: String(row[1] || ''), category: String(row[2] || ''), itemsCount: Number(row[3]) || 0, totalSum: Number(row[4]) || 0, sheetName: String(row[5] || '') });
  }
  out.reverse();
  return out;
}

function findEstimateById_(id) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(ESTIMATES_SHEET_NAME);
  if (!sh) return null;

  var values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][0]) === String(id)) {
      return { id: String(values[r][0]), name: String(values[r][1] || ''), category: String(values[r][2] || ''), itemsCount: Number(values[r][3]) || 0, totalSum: Number(values[r][4]) || 0, sheetName: String(values[r][5] || ''), _rowIndex: r + 1 };
    }
  }
  return null;
}

function updateEstimateTotals_(id, itemsCount, totalSum, sheetName) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(ESTIMATES_SHEET_NAME);
  var values = sh.getDataRange().getValues();

  for (var r = 1; r < values.length; r++) {
    if (String(values[r][0]) === String(id)) {
      sh.getRange(r + 1, 4, 1, 4).setValues([[itemsCount, totalSum, sheetName, new Date()]]);
      sh.getRange(r + 1, 5).setNumberFormat('0.00');
      return;
    }
  }
}

function buildItemsSheet_(sh, name, category) {
  sh.clear();
  sh.setFrozenRows(2);

  sh.getRange('A1:H1').merge();
  sh.getRange('A1').setValue(category + ' — ' + name).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#e5e7eb');

  sh.getRange(2, 1, 1, 8).setValues([['Статья', 'Кол-во', 'Залов', 'Дней', 'Дней мероприятий', 'Коэф.', 'Стоимость за ед.', 'Сумма']]);
  sh.getRange(2, 1, 1, 8).setFontWeight('bold').setBackground('#f3f4f6').setHorizontalAlignment('center');

  sh.setColumnWidth(1, 260);
  sh.setColumnWidths(2, 6, 120);
  sh.setColumnWidth(8, 140);

  sh.getRange(3, 1, 1, 8).setValues([['', 1, 1, 1, 1, 1, 0, '']]);
  sh.getRange(3, 8).setFormula('=B3*C3*D3*E3*F3*G3');

  sh.getRange('B3:F').setNumberFormat('0.########');
  sh.getRange('G3:H').setNumberFormat('0.00');

  applyArticleValidation_(sh, category, 1);
}

function applyArticleValidation_(itemsSheet, category, rowsCount) {
  var a = readRef_();
  var list = (a.articlesByCategory[category] || []).slice(0, 500);
  if (!list.length) return;

  var n = Math.max(1, rowsCount);
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
  itemsSheet.getRange(3, 1, n, 1).setDataValidation(rule);
}

function makeUniqueSheetName_(baseName, ss) {
  var cleaned = sanitizeSheetName_(baseName);
  var MAX = 99;

  var name = cleaned.length > MAX ? cleaned.slice(0, MAX) : cleaned;
  if (!ss.getSheetByName(name)) return name;

  for (var i = 2; i < 1000; i++) {
    var suffix = ' (' + i + ')';
    var candidate = name;
    if (candidate.length + suffix.length > MAX) candidate = candidate.slice(0, MAX - suffix.length);
    candidate += suffix;
    if (!ss.getSheetByName(candidate)) return candidate;
  }
  throw new Error('Не удалось подобрать уникальное имя листа.');
}

function sanitizeSheetName_(s) {
  return String(s || '').replace(/[\[\]\*\/\\\?\:]/g, ' ').replace(/\s+/g, ' ').trim();
}

function toNumber_(v, def) {
  if (def === undefined) def = 0;
  if (v === '' || v === null || v === undefined) return def;
  var n = Number(String(v).replace(',', '.'));
  return isFinite(n) ? n : def;
}
