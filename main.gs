var ESTIMATES_SHEET_NAME = 'Сметы';
var REF_SHEET_NAME = 'Справочник сметы';
var ITEMS_SHEET_PREFIX = 'Смета';
var ESTIMATES_STORE_KEY = 'estimates_store_v1';
var ITEMS_STORE_PREFIX = 'estimate_items_v1_';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Смета')
    .addItem('Конструктор сметы', 'openEstimateBuilder')
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
  var est = findEstimateById_(id);
  if (!est) throw new Error('Смета не найдена.');
  var template = HtmlService.createTemplateFromFile('EstimateItemsModal');
  template.estimateId = id;
  var html = template.evaluate()
    .setTitle('Позиции сметы')
    .setWidth(1100)
    .setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'Позиции сметы');
  return true;
}

function getBootstrapData() {
  ensureCoreSheets_();
  var a = readRef_();
  return { categories: a.categories, articlesByCategory: a.articlesByCategory, types: a.types, estimates: listEstimates_() };
}

function createEstimate(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();

    var name = String(payload && payload.name ? payload.name : '').trim();
    var category = String(payload && payload.category ? payload.category : '').trim();
    var type = String(payload && payload.type ? payload.type : '').trim();
    if (!name) throw new Error('Укажите название.');
    if (!category) throw new Error('Укажите категорию.');
    if (!type) throw new Error('Укажите тип.');

    var list = loadEstimates_();
    for (var i = 0; i < list.length; i++) {
      if (String(list[i].name || '').trim() === name && String(list[i].category || '').trim() === category) {
        return list[i];
      }
    }

    var id = Utilities.getUuid();
    var est = { id: id, name: name, category: category, type: type, itemsCount: 0, totalSum: 0, sheetName: '', updatedAt: new Date().toISOString() };
    list.push(est);
    saveEstimates_(list);
    return est;
  } finally {
    lock.releaseLock();
  }
}

function deleteEstimate(id) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();

    var list = loadEstimates_();
    for (var i = 0; i < list.length; i++) {
      if (String(list[i].id) === String(id)) {
        var sheetName = String(list[i].sheetName || '').trim();
        list.splice(i, 1);
        saveEstimates_(list);
        deleteItems_(id);
        if (sheetName) {
          var ss = SpreadsheetApp.getActive();
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
  return { estimate: est, items: loadItems_(id) };
}

function getEstimateItemsBootstrap(id) {
  ensureCoreSheets_();
  var est = findEstimateById_(id);
  if (!est) throw new Error('Смета не найдена.');
  var ref = readRef_();
  return { estimate: est, items: loadItems_(id), articlesByCategory: ref.articlesByCategory };
}

function saveEstimateItems(id, items) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();

    var est = findEstimateById_(id);
    if (!est) throw new Error('Смета не найдена.');

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

      out.push({ article: article, qty: qty, halls: halls, days: days, eventDays: eventDays, coef: coef, unitCost: unitCost });
      total += qty * halls * days * eventDays * coef * unitCost;
      count++;
    }

    saveItems_(id, out);
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

function exportEstimateToSheet(id) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();

    var est = findEstimateById_(id);
    if (!est) throw new Error('Смета не найдена.');

    var ss = SpreadsheetApp.getActive();
    var sheetName = String(est.sheetName || '').trim();
    var sh = sheetName ? ss.getSheetByName(sheetName) : null;
    if (!sh) {
      var baseSheetName = ITEMS_SHEET_PREFIX + ' • ' + est.category + ' • ' + est.name;
      sheetName = makeUniqueSheetName_(baseSheetName, ss);
      sh = ss.insertSheet(sheetName);
    }

    buildItemsSheet_(sh, est.name, est.category);

    var items = loadItems_(id);
    if (items.length) {
      var values = [];
      for (var i = 0; i < items.length; i++) {
        var it = items[i] || {};
        values.push([it.article || '', toNumber_(it.qty, 0), toNumber_(it.halls, 0), toNumber_(it.days, 0), toNumber_(it.eventDays, 0), toNumber_(it.coef, 1), toNumber_(it.unitCost, 0)]);
      }
      sh.getRange(3, 1, values.length, 7).setValues(values);

      var formulas = [];
      for (var k = 0; k < values.length; k++) {
        var rr = 3 + k;
        formulas.push(['=B' + rr + '*C' + rr + '*D' + rr + '*E' + rr + '*F' + rr + '*G' + rr]);
      }
      sh.getRange(3, 8, formulas.length, 1).setFormulas(formulas);

      sh.getRange(3, 2, values.length, 5).setNumberFormat('0.########');
      sh.getRange(3, 7, values.length, 2).setNumberFormat('0.00');
      applyArticleValidation_(sh, est.category, values.length);
    } else {
      applyArticleValidation_(sh, est.category, 1);
    }

    updateEstimateTotals_(id, est.itemsCount, est.totalSum, sheetName);
    ss.setActiveSheet(sh);
    return { sheetName: sheetName };
  } finally {
    lock.releaseLock();
  }
}

function ensureCoreSheets_() {
  var ss = SpreadsheetApp.getActive();

  if (!ss.getSheetByName(REF_SHEET_NAME)) {
    var ref = ss.insertSheet(REF_SHEET_NAME);
    ref.getRange('A1:B1').setValues([['Категория', 'Статья']]);
    ref.getRange('E1').setValue('Тип');
    ref.setFrozenRows(1);
    ref.autoResizeColumns(1, 2);
  }
}

function readRef_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF_SHEET_NAME);
  if (!sh) return { categories: [], articlesByCategory: {}, types: [] };

  var last = sh.getLastRow();
  if (last < 2) return { categories: [], articlesByCategory: {}, types: [] };

  var cats = sh.getRange(2, 1, last - 1, 1).getValues();
  var arts = sh.getRange(2, 2, last - 1, 1).getValues();
  var typesRange = sh.getRange(2, 5, last - 1, 1).getValues();

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

  var typesSeen = {};
  var types = [];
  for (var t = 0; t < typesRange.length; t++) {
    var typeVal = String(typesRange[t][0] || '').trim();
    if (!typeVal || typesSeen[typeVal]) continue;
    typesSeen[typeVal] = true;
    types.push(typeVal);
  }
  types.sort(function(x, y) { return x.localeCompare(y, 'ru'); });

  return { categories: categories, articlesByCategory: map, types: types };
}

function listEstimates_() {
  var out = loadEstimates_();
  out.reverse();
  return out;
}

function findEstimateById_(id) {
  var list = loadEstimates_();
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].id) === String(id)) {
      return list[i];
    }
  }
  return null;
}

function updateEstimateTotals_(id, itemsCount, totalSum, sheetName) {
  var list = loadEstimates_();
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].id) === String(id)) {
      list[i].itemsCount = itemsCount;
      list[i].totalSum = totalSum;
      list[i].sheetName = sheetName || '';
      list[i].updatedAt = new Date().toISOString();
      saveEstimates_(list);
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

function loadEstimates_() {
  var props = PropertiesService.getDocumentProperties();
  var raw = props.getProperty(ESTIMATES_STORE_KEY);
  if (!raw) {
    var legacy = loadLegacyEstimatesFromSheet_();
    if (legacy.length) {
      saveEstimates_(legacy);
      return legacy;
    }
    return [];
  }
  try {
    var parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function saveEstimates_(list) {
  PropertiesService.getDocumentProperties().setProperty(ESTIMATES_STORE_KEY, JSON.stringify(list || []));
}

function loadLegacyEstimatesFromSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(ESTIMATES_SHEET_NAME);
  if (!sh) return [];

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(normalizeHeader_);
  var map = inferLegacyEstimatesMap_(headers, values);
  var startIndex = map.hasHeader ? 1 : 0;

  var out = [];
  for (var r = startIndex; r < values.length; r++) {
    var row = values[r] || [];
    var name = String(cellAt_(row, map.name) || '').trim();
    var category = String(cellAt_(row, map.category) || '').trim();
    if (!name && !category) continue;

    var id = String(cellAt_(row, map.id) || '').trim();
    if (!id) id = Utilities.getUuid();

    var itemsCount = toNumber_(cellAt_(row, map.itemsCount), 0);
    var totalSum = toNumber_(cellAt_(row, map.totalSum), 0);
    var sheetName = String(cellAt_(row, map.sheetName) || '').trim();
    var type = String(cellAt_(row, map.type) || '').trim();

    var updatedAtRaw = cellAt_(row, map.updatedAt);
    var updatedAt = '';
    if (updatedAtRaw instanceof Date) {
      updatedAt = updatedAtRaw.toISOString();
    } else {
      updatedAt = String(updatedAtRaw || '').trim();
    }
    if (!updatedAt) updatedAt = new Date().toISOString();

    out.push({
      id: id,
      name: name,
      category: category,
      type: type,
      itemsCount: itemsCount,
      totalSum: totalSum,
      sheetName: sheetName,
      updatedAt: updatedAt
    });
  }
  return out;
}

function inferLegacyEstimatesMap_(headers, values) {
  var idx = {
    id: findHeaderIndex_(headers, ['id', 'ид', 'uuid', 'guid']),
    name: findHeaderIndex_(headers, ['название', 'имя', 'наименование']),
    category: findHeaderIndex_(headers, ['категория', 'кат']),
    itemsCount: findHeaderIndex_(headers, ['статей', 'позиц', 'кол-во', 'количество', 'items']),
    totalSum: findHeaderIndex_(headers, ['сумма', 'итого', 'total']),
    sheetName: findHeaderIndex_(headers, ['лист', 'sheet']),
    type: findHeaderIndex_(headers, ['тип']),
    updatedAt: findHeaderIndex_(headers, ['обнов', 'дата', 'updated'])
  };

  var hasHeader = false;
  for (var key in idx) {
    if (idx[key] >= 0) {
      hasHeader = true;
      break;
    }
  }

  if (!hasHeader) {
    idx = {
      id: -1,
      name: 0,
      category: 1,
      itemsCount: 2,
      totalSum: 3,
      sheetName: 4,
      type: -1,
      updatedAt: 5
    };
  }

  if (idx.id < 0) {
    idx.id = detectUuidColumn_(values, hasHeader ? 1 : 0);
  }

  return {
    hasHeader: hasHeader,
    id: idx.id,
    name: idx.name,
    category: idx.category,
    itemsCount: idx.itemsCount,
    totalSum: idx.totalSum,
    sheetName: idx.sheetName,
    type: idx.type,
    updatedAt: idx.updatedAt
  };
}

function normalizeHeader_(value) {
  return String(value || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function findHeaderIndex_(headers, candidates) {
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (!h) continue;
    for (var j = 0; j < candidates.length; j++) {
      var c = candidates[j];
      if (h === c || h.indexOf(c) !== -1) return i;
    }
  }
  return -1;
}

function detectUuidColumn_(values, startIndex) {
  var uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  var sampleRows = Math.min(values.length, startIndex + 20);
  var maxCols = values[0] ? values[0].length : 0;
  for (var c = 0; c < maxCols; c++) {
    for (var r = startIndex; r < sampleRows; r++) {
      var cell = values[r] ? values[r][c] : '';
      if (typeof cell === 'string' && uuid.test(cell.trim())) return c;
    }
  }
  return -1;
}

function cellAt_(row, index) {
  if (index === undefined || index === null || index < 0) return '';
  return row[index];
}

function loadItems_(id) {
  var props = PropertiesService.getDocumentProperties();
  var raw = props.getProperty(ITEMS_STORE_PREFIX + id);
  if (!raw) return [];
  try {
    var parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function saveItems_(id, items) {
  PropertiesService.getDocumentProperties().setProperty(ITEMS_STORE_PREFIX + id, JSON.stringify(items || []));
}

function deleteItems_(id) {
  PropertiesService.getDocumentProperties().deleteProperty(ITEMS_STORE_PREFIX + id);
}
