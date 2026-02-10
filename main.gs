var REF_SHEET_NAME = 'Справочник сметы';
var CLIENTS_DB_SHEET_NAME = 'БД сметы';

var PROJECTS_STORE_KEY = 'cs_projects_v3';
var ENTRIES_STORE_PREFIX = 'cs_entries_v3_';
var ITEMS_STORE_PREFIX = 'cs_items_v3_';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Смета')
    .addItem('Конструктор сметы', 'openEstimateBuilder')
    .addToUi();
}

function openEstimateBuilder() {
  ensureCoreSheets_();
  var html = HtmlService.createHtmlOutputFromFile('EstimateModal')
    .setTitle('Конструктор сметы')
    .setWidth(1240)
    .setHeight(740);
  SpreadsheetApp.getUi().showModalDialog(html, 'Конструктор сметы');
}

function openEntryItems(entryId) {
  ensureCoreSheets_();
  var info = findEntryWithProject_(entryId);
  if (!info || !info.entry) throw new Error('Запись сметы не найдена.');

  var tpl = HtmlService.createTemplateFromFile('EstimateItemsModal');
  tpl.entryId = String(entryId);

  var html = tpl.evaluate()
    .setTitle('Позиции сметы')
    .setWidth(1180)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, 'Позиции сметы');
  return true;
}

function getBootstrapData() {
  ensureCoreSheets_();
  var ref = readRef_();
  var projects = listProjects_();
  return {
    projects: projects,
    categories: ref.categories,
    positionsByCategory: ref.positionsByCategory,
    types: ref.types,
    pricesByPosition: ref.pricesByPosition,
    clientsDbRows: readClientsDbRows_()
  };
}

function getProjectEntries(projectId) {
  ensureCoreSheets_();
  var project = findProjectById_(projectId);
  if (!project) throw new Error('Проект не найден.');
  var entries = loadEntries_(projectId);
  entries.sort(function(a, b) {
    var at = String(a.type || '').localeCompare(String(b.type || ''), 'ru');
    if (at) return at;
    var ag = String(a.group || '').localeCompare(String(b.group || ''), 'ru');
    if (ag) return ag;
    return String(a.category || '').localeCompare(String(b.category || ''), 'ru');
  });
  return { project: project, entries: entries };
}

function exportProjectToDraft(projectId) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();
    var project = findProjectById_(projectId);
    if (!project) throw new Error('Проект не найден.');

    var entries = loadEntries_(projectId);
    if (!entries.length) return { ok: true, rows: 0 };

    var rows = [];
    var exportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    var client = String(project.client || '').trim();
    var projectName = String(project.name || '').trim();

    for (var e = 0; e < entries.length; e++) {
      var entry = entries[e] || {};
      var items = normalizeItems_(loadItems_(String(entry.id)));
      for (var i = 0; i < items.length; i++) {
        var it = items[i] || {};
        rows.push([
          exportDate,
          client,
          projectName,
          String(entry.type || ''),
          String(entry.group || ''),
          String(entry.category || ''),
          String(it.position || ''),
          toNumber_(it.qty, 0),
          toNumber_(it.halls, 0),
          toNumber_(it.days, 0),
          toNumber_(it.coef, 1),
          toNumber_(it.unitCost, 0)
        ]);
      }
    }

    if (!rows.length) return { ok: true, rows: 0 };

    var sh = getOrCreateDraftSheet_();
    var startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, rows.length, 12).setValues(rows);
    return { ok: true, rows: rows.length };
  } finally {
    lock.releaseLock();
  }
}

function exportProjectToClientSheet(projectId) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();
    var project = findProjectById_(projectId);
    if (!project) throw new Error('Проект не найден.');

    var entries = loadEntries_(projectId);
    if (!entries.length) return { ok: true, rows: 0, sheetName: '' };

    entries.sort(function(a, b) {
      var ag = String(a.group || '').localeCompare(String(b.group || ''), 'ru');
      if (ag) return ag;
      var ac = String(a.category || '').localeCompare(String(b.category || ''), 'ru');
      if (ac) return ac;
      return String(a.type || '').localeCompare(String(b.type || ''), 'ru');
    });

    var client = String(project.client || '').trim();
    var projectName = String(project.name || '').trim();
    var sheetName = makeUniqueSheetName_(buildClientProjectSheetName_(client, projectName));

    var ss = SpreadsheetApp.getActive();
    var sh = ss.insertSheet(sheetName);

    sh.setHiddenGridlines(false);
    sh.setColumnWidths(2, 1, 420); // B
    sh.setColumnWidths(3, 4, 56);  // C-F
    sh.setColumnWidths(7, 2, 120); // G-H
    sh.setColumnWidths(9, 1, 620); // I

    sh.getRange('B1:H3').merge();
    sh.getRange('B1').setValue((client ? client + ' — ' : '') + projectName);
    sh.getRange('B1:H3')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setFontWeight('bold')
      .setFontSize(18)
      .setFontFamily('Arial');

    sh.getRange(4, 2, 1, 8).setValues([['Наименование позиции', 'Кол-во', 'Залов', 'Дней', 'Коэф.', 'Стоимость за ед.', 'Итоговая стоимость', 'Комментарии']]);
    sh.getRange(4, 2, 1, 8)
      .setFontWeight('bold')
      .setFontFamily('Arial')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#d0d0d0', SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(5, 2, 1, 8).merge();
    sh.getRange(5, 2)
      .setValue(String(projectName || 'Проект'))
      .setFontWeight('bold')
      .setFontSize(22)
      .setFontFamily('Arial')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBackground('#3f3f3f')
      .setFontColor('#ffffff');
    sh.setRowHeight(5, 40);

    var row = 6;
    var rowsCount = 0;
    var currentGroup = '';
    var currentCategory = '';

    for (var e = 0; e < entries.length; e++) {
      var entry = entries[e] || {};
      var groupLabel = String(entry.group || '').trim();
      var categoryLabel = String(entry.category || '').trim();

      if (groupLabel && groupLabel !== currentGroup) {
        sh.getRange(row, 2, 1, 8).merge();
        sh.getRange(row, 2)
          .setValue(groupLabel)
          .setFontWeight('bold')
          .setFontSize(24)
          .setFontFamily('Arial')
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
        sh.setRowHeight(row, 34);
        row++;
        currentGroup = groupLabel;
        currentCategory = '';
      }

      if (categoryLabel && categoryLabel !== currentCategory) {
        sh.getRange(row, 2, 1, 8).merge();
        sh.getRange(row, 2)
          .setValue(categoryLabel)
          .setFontStyle('italic')
          .setFontColor('#9ca3af')
          .setFontSize(16)
          .setFontFamily('Arial')
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
        sh.setRowHeight(row, 26);
        row++;
        currentCategory = categoryLabel;
      }

      var items = normalizeItems_(loadItems_(String(entry.id)));
      for (var i = 0; i < items.length; i++) {
        var it = items[i] || {};
        var lineTotal = toNumber_(it.qty, 0) * toNumber_(it.halls, 0) * toNumber_(it.days, 0) * toNumber_(it.eventDays, 1) * toNumber_(it.coef, 1) * toNumber_(it.unitCost, 0);
        sh.getRange(row, 2, 1, 8).setValues([[String(it.position || ''), toNumber_(it.qty, 0), toNumber_(it.halls, 0), toNumber_(it.days, 0), toNumber_(it.coef, 1), toNumber_(it.unitCost, 0), lineTotal, String(it.comment || '')]]);
        sh.setRowHeight(row, 29);
        row++;
        rowsCount++;
      }
    }

    var dataRows = Math.max(row - 6, 1);
    sh.getRange(6, 2, dataRows, 8)
      .setFontFamily('Arial')
      .setFontSize(11)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#d0d0d0', SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(6, 3, dataRows, 5).setHorizontalAlignment('center');
    sh.getRange(6, 7, dataRows, 2).setNumberFormat('#,##0"₽"');
    sh.getRange(6, 2, dataRows, 1).setHorizontalAlignment('left');
    sh.getRange(6, 9, dataRows, 1).setHorizontalAlignment('left');

    sh.setFrozenRows(5);

    return { ok: true, rows: rowsCount, sheetName: sheetName };
  } finally {
    lock.releaseLock();
  }
}

function createProject(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();
    var name = String(payload && payload.name ? payload.name : '').trim();
    if (!name) throw new Error('Укажите наименование проекта.');

    var client = String(payload && payload.client ? payload.client : '').trim();
    var tariff = String(payload && payload.tariff ? payload.tariff : '').trim();
    if (!tariff) tariff = 'Обычный';

    var list = loadProjects_();
    var id = Utilities.getUuid();
    var p = { id: id, name: name, client: client, tariff: tariff, itemsCount: 0, totalSum: 0, updatedAt: new Date().toISOString() };
    list.push(p);
    saveProjects_(list);
    saveEntries_(id, []);
    return p;
  } finally {
    lock.releaseLock();
  }
}

function importExistingProject(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();

    var name = String(payload && payload.name ? payload.name : '').trim();
    var tariff = String(payload && payload.tariff ? payload.tariff : '').trim();
    var client = String(payload && payload.client ? payload.client : '').trim();
    var project = String(payload && payload.project ? payload.project : '').trim();
    var date = String(payload && payload.date ? payload.date : '').trim();

    if (!name) throw new Error('Укажите наименование проекта.');
    if (!tariff) tariff = 'Обычный';
    if (!client || !project || !date) throw new Error('Не выбраны данные клиента для импорта.');

    var rows = readClientsDbRows_().filter(function(r) {
      return String(r.client) === client && String(r.project) === project && String(r.date) === date;
    });
    if (!rows.length) throw new Error('По выбранным клиенту/проекту/дате данные не найдены.');

    var projects = loadProjects_();
    var projectId = Utilities.getUuid();
    var newProject = {
      id: projectId,
      name: name,
      client: client,
      tariff: tariff,
      itemsCount: 0,
      totalSum: 0,
      updatedAt: new Date().toISOString()
    };
    projects.push(newProject);
    saveProjects_(projects);

    var groupsMap = {};
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i] || {};
      var type = String(row.type || '').trim();
      var group = String(row.group || '').trim();
      var category = String(row.category || '').trim();
      var position = String(row.position || '').trim();
      if (!type || !category || !position) continue;

      var key = [type, group, category].join('|||');
      if (!groupsMap[key]) {
        groupsMap[key] = {
          id: Utilities.getUuid(),
          projectId: projectId,
          type: type,
          group: group,
          category: category,
          items: []
        };
      }

      groupsMap[key].items.push({
        position: position,
        qty: toNumber_(row.qty, 0),
        halls: toNumber_(row.halls, 0),
        days: toNumber_(row.days, 0),
        eventDays: 1,
        coef: toNumber_(row.coef, 1),
        unitCost: toNumber_(row.unitCost, 0),
        comment: ''
      });
    }

    var entries = [];
    var keys = Object.keys(groupsMap);
    for (var k = 0; k < keys.length; k++) {
      var g = groupsMap[keys[k]];
      saveItems_(g.id, g.items);
      var totals = computeTotals_(g.items);
      entries.push({
        id: g.id,
        projectId: projectId,
        type: g.type,
        group: g.group,
        category: g.category,
        itemsCount: totals.itemsCount,
        totalSum: totals.totalSum,
        updatedAt: new Date().toISOString()
      });
    }

    saveEntries_(projectId, entries);
    recalcProjectTotals_(projectId);

    return findProjectById_(projectId);
  } finally {
    lock.releaseLock();
  }
}

function deleteProject(projectId) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();

    var list = loadProjects_();
    var idx = -1;
    for (var i = 0; i < list.length; i++) {
      if (String(list[i].id) === String(projectId)) { idx = i; break; }
    }
    if (idx < 0) return false;

    var entries = loadEntries_(projectId);
    for (var e = 0; e < entries.length; e++) {
      var entryId = String(entries[e].id);
      PropertiesService.getDocumentProperties().deleteProperty(ITEMS_STORE_PREFIX + entryId);
    }
    PropertiesService.getDocumentProperties().deleteProperty(ENTRIES_STORE_PREFIX + projectId);

    list.splice(idx, 1);
    saveProjects_(list);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function duplicateProject(projectId) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();

    var src = findProjectById_(projectId);
    if (!src) throw new Error('Проект не найден.');

    var newId = Utilities.getUuid();
    var list = loadProjects_();
    var p = { id: newId, name: String(src.name || '').trim() + ' (копия)', client: String(src.client || ''), tariff: String(src.tariff || 'Обычный'), itemsCount: 0, totalSum: 0, updatedAt: new Date().toISOString() };
    list.push(p);
    saveProjects_(list);

    var srcEntries = loadEntries_(projectId);
    var newEntries = [];
    for (var i = 0; i < srcEntries.length; i++) {
      var se = srcEntries[i] || {};
      var neId = Utilities.getUuid();
      var ne = {
        id: neId,
        projectId: newId,
        type: String(se.type || ''),
        group: String(se.group || ''),
        category: String(se.category || ''),
        itemsCount: 0,
        totalSum: 0,
        updatedAt: new Date().toISOString()
      };
      newEntries.push(ne);

      var srcItems = loadItems_(String(se.id));
      saveItems_(neId, srcItems);

      var totals = computeTotals_(srcItems);
      ne.itemsCount = totals.itemsCount;
      ne.totalSum = totals.totalSum;
    }
    saveEntries_(newId, newEntries);

    recalcProjectTotals_(newId);
    return p;
  } finally {
    lock.releaseLock();
  }
}

function createEntry(projectId, payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();
    var project = findProjectById_(projectId);
    if (!project) throw new Error('Проект не найден.');

    var type = String(payload && payload.type ? payload.type : '').trim();
    var group = String(payload && payload.group ? payload.group : '').trim();
    var category = String(payload && payload.category ? payload.category : '').trim();

    if (!type) throw new Error('Укажите тип.');
    if (!category) throw new Error('Укажите категорию.');

    var entries = loadEntries_(projectId);
    var id = Utilities.getUuid();
    var e = {
      id: id,
      projectId: String(projectId),
      type: type,
      group: group,
      category: category,
      itemsCount: 0,
      totalSum: 0,
      updatedAt: new Date().toISOString()
    };
    entries.push(e);
    saveEntries_(projectId, entries);
    saveItems_(id, []);
    recalcProjectTotals_(projectId);
    return e;
  } finally {
    lock.releaseLock();
  }
}

function deleteEntry(entryId) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();
    var info = findEntryWithProject_(entryId);
    if (!info || !info.project || !info.entry) return false;

    var projectId = String(info.project.id);
    var entries = loadEntries_(projectId);
    var out = [];
    for (var i = 0; i < entries.length; i++) {
      if (String(entries[i].id) !== String(entryId)) out.push(entries[i]);
    }
    saveEntries_(projectId, out);
    PropertiesService.getDocumentProperties().deleteProperty(ITEMS_STORE_PREFIX + entryId);
    recalcProjectTotals_(projectId);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function duplicateEntry(entryId) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();
    var info = findEntryWithProject_(entryId);
    if (!info || !info.project || !info.entry) throw new Error('Запись сметы не найдена.');

    var projectId = String(info.project.id);
    var src = info.entry;

    var newId = Utilities.getUuid();
    var ne = {
      id: newId,
      projectId: projectId,
      type: String(src.type || ''),
      group: String(src.group || '') + ' (копия)',
      category: String(src.category || ''),
      itemsCount: 0,
      totalSum: 0,
      updatedAt: new Date().toISOString()
    };

    var items = loadItems_(entryId);
    saveItems_(newId, items);
    var totals = computeTotals_(items);
    ne.itemsCount = totals.itemsCount;
    ne.totalSum = totals.totalSum;

    var entries = loadEntries_(projectId);
    entries.push(ne);
    saveEntries_(projectId, entries);
    recalcProjectTotals_(projectId);
    return ne;
  } finally {
    lock.releaseLock();
  }
}

function getEntryItemsBootstrap(entryId) {
  ensureCoreSheets_();
  var info = findEntryWithProject_(entryId);
  if (!info || !info.project || !info.entry) throw new Error('Запись сметы не найдена.');

  var ref = readRef_();
  var items = normalizeItems_(loadItems_(entryId));
  return {
    project: info.project,
    entry: info.entry,
    items: items,
    categories: ref.categories,
    positionsByCategory: ref.positionsByCategory,
    types: ref.types,
    pricesByPosition: ref.pricesByPosition,
    clientsDbRows: readClientsDbRows_()
  };
}

function saveEntryAll(entryId, meta, items) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();
    var info = findEntryWithProject_(entryId);
    if (!info || !info.project || !info.entry) throw new Error('Запись сметы не найдена.');

    var projectId = String(info.project.id);

    var projectName = String(meta && meta.projectName ? meta.projectName : info.project.name).trim();
    var type = String(meta && meta.type ? meta.type : info.entry.type).trim();
    var group = String(meta && meta.group ? meta.group : info.entry.group).trim();
    var category = String(meta && meta.category ? meta.category : info.entry.category).trim();

    if (!projectName) throw new Error('Укажите наименование проекта.');
    if (!type) throw new Error('Укажите тип.');
    if (!category) throw new Error('Укажите категорию.');

    updateProjectName_(projectId, projectName);

    var cleaned = [];
    var src = Array.isArray(items) ? items : [];
    for (var i = 0; i < src.length; i++) {
      var it = src[i] || {};
      var position = String(it.position || '').trim();
      if (!position) continue;

      cleaned.push({
        position: position,
        qty: toNumber_(it.qty, 0),
        halls: toNumber_(it.halls, 0),
        days: toNumber_(it.days, 0),
        eventDays: toNumber_(it.eventDays, 0),
        coef: toNumber_(it.coef, 1),
        unitCost: toNumber_(it.unitCost, 0),
        comment: String(it.comment || '')
      });
    }

    saveItems_(entryId, cleaned);

    var totals = computeTotals_(cleaned);

    var entries = loadEntries_(projectId);
    for (var e = 0; e < entries.length; e++) {
      if (String(entries[e].id) === String(entryId)) {
        entries[e].type = type;
        entries[e].group = group;
        entries[e].category = category;
        entries[e].itemsCount = totals.itemsCount;
        entries[e].totalSum = totals.totalSum;
        entries[e].updatedAt = new Date().toISOString();
        break;
      }
    }
    saveEntries_(projectId, entries);

    recalcProjectTotals_(projectId);

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateDraftSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Черновик');
  if (!sh) sh = ss.insertSheet('Черновик');

  var headers = ['Дата', 'Клиент', 'Проект', 'Тип', 'Группа', 'Категория', 'Позиция', 'Кол-во', 'Залов', 'Дней', 'Коэф.', 'Цена'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  return sh;
}

function sanitizeSheetPart_(value) {
  var out = String(value || '').trim();
  out = out.replace(/[\\\/\?\*\[\]:]/g, '_');
  out = out.replace(/\s+/g, ' ');
  return out;
}

function buildClientProjectSheetName_(client, projectName) {
  var c = sanitizeSheetPart_(client) || 'Клиент';
  var p = sanitizeSheetPart_(projectName) || 'Проект';
  var base = c + '_' + p;
  if (base.length > 95) base = base.slice(0, 95);
  return base;
}

function makeUniqueSheetName_(baseName) {
  var ss = SpreadsheetApp.getActive();
  var base = sanitizeSheetPart_(baseName) || 'Лист экспорта';
  if (base.length > 95) base = base.slice(0, 95);

  var name = base;
  var n = 2;
  while (ss.getSheetByName(name)) {
    var suffix = ' (' + n + ')';
    var maxBase = 100 - suffix.length;
    var shortBase = base.length > maxBase ? base.slice(0, maxBase) : base;
    name = shortBase + suffix;
    n++;
  }
  return name;
}

function ensureCoreSheets_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(REF_SHEET_NAME)) {
    var ref = ss.insertSheet(REF_SHEET_NAME);
    ref.getRange('A1').setValue('Категория');
    ref.getRange('B1').setValue('Позиция');
    ref.getRange('D1').setValue('Тип');
    ref.setFrozenRows(1);
    ref.autoResizeColumns(1, 4);
  }
}

function readRef_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REF_SHEET_NAME);
  if (!sh) return { categories: [], positionsByCategory: {}, types: [], pricesByPosition: {} };

  var last = sh.getLastRow();
  if (last < 2) return { categories: [], positionsByCategory: {}, types: [], pricesByPosition: {} };

  var cats = sh.getRange(2, 1, last - 1, 1).getValues();
  var pos = sh.getRange(2, 2, last - 1, 1).getValues();
  var prices = sh.getRange(2, 3, last - 1, 1).getValues();
  var typesRange = sh.getRange(2, 4, last - 1, 1).getValues();

  var map = {};
  var pricesMap = {};
  for (var i = 0; i < cats.length; i++) {
    var c = String(cats[i][0] || '').trim();
    var p = String(pos[i][0] || '').trim();
    var price = toNumber_(prices[i][0], 0);
    if (!c || !p) continue;
    if (!map[c]) map[c] = [];
    map[c].push(p);
    if (price || price === 0) pricesMap[p] = price;
  }

  var categories = Object.keys(map).sort(function(x, y) { return x.localeCompare(y, 'ru'); });
  for (var j = 0; j < categories.length; j++) {
    var key = categories[j];
    var seen = {};
    var uniq = [];
    var a = map[key] || [];
    for (var k = 0; k < a.length; k++) {
      var v = String(a[k] || '').trim();
      if (!v || seen[v]) continue;
      seen[v] = true;
      uniq.push(v);
    }
    uniq.sort(function(x, y) { return x.localeCompare(y, 'ru'); });
    map[key] = uniq;
  }

  var typesSeen = {};
  var types = [];
  for (var t = 0; t < typesRange.length; t++) {
    var tv = String(typesRange[t][0] || '').trim();
    if (!tv || typesSeen[tv]) continue;
    typesSeen[tv] = true;
    types.push(tv);
  }
  types.sort(function(x, y) { return x.localeCompare(y, 'ru'); });

  return { categories: categories, positionsByCategory: map, types: types, pricesByPosition: pricesMap };
}

function readClientsDbRows_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CLIENTS_DB_SHEET_NAME);
  if (!sh) return [];

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 12) return [];

  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  var headerMap = {};
  for (var h = 0; h < headers.length; h++) {
    var key = String(headers[h] || '').trim().toLowerCase();
    if (!key) continue;
    headerMap[key] = h;
  }

  function colIndexByAliases_(aliases, fallbackIndex) {
    for (var a = 0; a < aliases.length; a++) {
      var k = String(aliases[a] || '').trim().toLowerCase();
      if (Object.prototype.hasOwnProperty.call(headerMap, k)) return headerMap[k];
    }
    return fallbackIndex;
  }

  var idxDate = colIndexByAliases_(['дата'], 0);
  var idxClient = colIndexByAliases_(['клиент'], 1);
  var idxProject = colIndexByAliases_(['проект'], 2);
  var idxType = colIndexByAliases_(['тип'], 3);
  var idxGroup = colIndexByAliases_(['группа'], 4);
  var idxCategory = colIndexByAliases_(['категория'], 5);
  var idxPosition = colIndexByAliases_(['позиция', 'наименование'], 6);
  var idxQty = colIndexByAliases_(['кол-во', 'колво', 'количество'], 7);
  var idxHalls = colIndexByAliases_(['залов', 'залы'], 8);
  var idxDays = colIndexByAliases_(['дней', 'дни'], 9);
  var idxCoef = colIndexByAliases_(['коэф.', 'коэф', 'коэффициент'], 10);
  var idxPrice = colIndexByAliases_(['цена', 'стоимость'], 11);

  var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i] || [];

    var dateRaw = r[idxDate];
    var date = '';
    if (Object.prototype.toString.call(dateRaw) === '[object Date]' && !isNaN(dateRaw.getTime())) {
      date = Utilities.formatDate(dateRaw, Session.getScriptTimeZone(), 'dd.MM.yyyy');
    } else {
      date = String(dateRaw || '').trim();
    }

    var client = String(r[idxClient] || '').trim();
    var project = String(r[idxProject] || '').trim();
    var type = String(r[idxType] || '').trim();
    var group = String(r[idxGroup] || '').trim();
    var category = String(r[idxCategory] || '').trim();
    var position = String(r[idxPosition] || '').trim();

    if (!date || !client || !project || !type || !position) continue;

    out.push({
      date: date,
      client: client,
      project: project,
      category: category,
      type: type,
      group: group,
      position: position,
      qty: toNumber_(r[idxQty], 0),
      halls: toNumber_(r[idxHalls], 0),
      days: toNumber_(r[idxDays], 0),
      coef: toNumber_(r[idxCoef], 1),
      unitCost: toNumber_(r[idxPrice], 0)
    });
  }

  return out;
}

function listProjects_() {
  var list = loadProjects_();
  list.sort(function(a, b) { return String(a.name || '').localeCompare(String(b.name || ''), 'ru'); });
  return list;
}

function findProjectById_(id) {
  var list = loadProjects_();
  for (var i = 0; i < list.length; i++) if (String(list[i].id) === String(id)) return list[i];
  return null;
}

function updateProjectName_(projectId, newName) {
  var list = loadProjects_();
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].id) === String(projectId)) {
      list[i].name = String(newName || '').trim();
      list[i].updatedAt = new Date().toISOString();
      saveProjects_(list);
      return;
    }
  }
}

function recalcProjectTotals_(projectId) {
  var entries = loadEntries_(projectId);
  var itemsCount = 0;
  var totalSum = 0;
  for (var i = 0; i < entries.length; i++) {
    itemsCount += Number(entries[i].itemsCount || 0);
    totalSum += Number(entries[i].totalSum || 0);
  }

  var list = loadProjects_();
  for (var j = 0; j < list.length; j++) {
    if (String(list[j].id) === String(projectId)) {
      list[j].itemsCount = itemsCount;
      list[j].totalSum = totalSum;
      list[j].updatedAt = new Date().toISOString();
      saveProjects_(list);
      return;
    }
  }
}

function findEntryWithProject_(entryId) {
  var projects = loadProjects_();
  for (var p = 0; p < projects.length; p++) {
    var pid = String(projects[p].id);
    var entries = loadEntries_(pid);
    for (var e = 0; e < entries.length; e++) {
      if (String(entries[e].id) === String(entryId)) {
        return { project: projects[p], entry: entries[e] };
      }
    }
  }
  return null;
}

function normalizeItems_(items) {
  var src = Array.isArray(items) ? items : [];
  var out = [];
  for (var i = 0; i < src.length; i++) {
    var it = src[i] || {};
    var position = String(it.position || '').trim();
    if (!position) continue;
    out.push({
      position: position,
      qty: toNumber_(it.qty, 0),
      halls: toNumber_(it.halls, 0),
      days: toNumber_(it.days, 0),
      eventDays: toNumber_(it.eventDays, 0),
      coef: toNumber_(it.coef, 1),
      unitCost: toNumber_(it.unitCost, 0),
      comment: String(it.comment || '')
    });
  }
  return out;
}

function computeTotals_(items) {
  var src = Array.isArray(items) ? items : [];
  var total = 0;
  var count = 0;
  for (var i = 0; i < src.length; i++) {
    var it = src[i] || {};
    var position = String(it.position || '').trim();
    if (!position) continue;
    total += toNumber_(it.qty, 0) * toNumber_(it.halls, 0) * toNumber_(it.days, 0) * toNumber_(it.eventDays, 0) * toNumber_(it.coef, 1) * toNumber_(it.unitCost, 0);
    count++;
  }
  return { itemsCount: count, totalSum: total };
}

function toNumber_(v, def) {
  if (def === undefined) def = 0;
  if (v === '' || v === null || v === undefined) return def;
  var n = Number(String(v).replace(',', '.'));
  return isFinite(n) ? n : def;
}

function loadProjects_() {
  var raw = PropertiesService.getDocumentProperties().getProperty(PROJECTS_STORE_KEY);
  if (!raw) return [];
  try {
    var parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function saveProjects_(list) {
  PropertiesService.getDocumentProperties().setProperty(PROJECTS_STORE_KEY, JSON.stringify(list || []));
}

function loadEntries_(projectId) {
  var raw = PropertiesService.getDocumentProperties().getProperty(ENTRIES_STORE_PREFIX + projectId);
  if (!raw) return [];
  try {
    var parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function saveEntries_(projectId, entries) {
  PropertiesService.getDocumentProperties().setProperty(ENTRIES_STORE_PREFIX + projectId, JSON.stringify(entries || []));
}

function loadItems_(entryId) {
  var raw = PropertiesService.getDocumentProperties().getProperty(ITEMS_STORE_PREFIX + entryId);
  if (!raw) return [];
  try {
    var parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function saveItems_(entryId, items) {
  PropertiesService.getDocumentProperties().setProperty(ITEMS_STORE_PREFIX + entryId, JSON.stringify(items || []));
}
