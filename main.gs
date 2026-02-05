var REF_SHEET_NAME = 'Справочник сметы';

var PROJECTS_STORE_KEY = 'cs_projects_v3';
var ENTRIES_STORE_PREFIX = 'cs_entries_v3_';        // + projectId
var ITEMS_STORE_PREFIX = 'cs_items_v3_';            // + entryId

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

/** BOOTSTRAP (первое окно) */
function getBootstrapData() {
  ensureCoreSheets_();
  var ref = readRef_();
  var projects = listProjects_();
  return { projects: projects, categories: ref.categories, positionsByCategory: ref.positionsByCategory, types: ref.types, pricesByPosition: ref.pricesByPosition };
}

/** ПРАВЫЙ БЛОК: список записей сметы по проекту */
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

/** ПРОЕКТЫ */
function createProject(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    ensureCoreSheets_();
    var name = String(payload && payload.name ? payload.name : '').trim();
    if (!name) throw new Error('Укажите наименование проекта.');

    var list = loadProjects_();
    var id = Utilities.getUuid();
    var p = { id: id, name: name, itemsCount: 0, totalSum: 0, updatedAt: new Date().toISOString() };
    list.push(p);
    saveProjects_(list);
    saveEntries_(id, []);
    return p;
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
    var p = { id: newId, name: String(src.name || '').trim() + ' (копия)', itemsCount: 0, totalSum: 0, updatedAt: new Date().toISOString() };
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

/** ЗАПИСИ (Смета) */
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

/** BOOTSTRAP (второе окно) */
function getEntryItemsBootstrap(entryId) {
  ensureCoreSheets_();
  var info = findEntryWithProject_(entryId);
  if (!info || !info.project || !info.entry) throw new Error('Запись сметы не найдена.');

  var ref = readRef_();
  var items = normalizeItems_(loadItems_(entryId));
  return { project: info.project, entry: info.entry, items: items, categories: ref.categories, positionsByCategory: ref.positionsByCategory, types: ref.types, pricesByPosition: ref.pricesByPosition };
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

/** HELPERS */
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

  var cats = sh.getRange(2, 1, last - 1, 1).getValues();  // A
  var pos = sh.getRange(2, 2, last - 1, 1).getValues();   // B
  var prices = sh.getRange(2, 3, last - 1, 1).getValues(); // C
  var typesRange = sh.getRange(2, 4, last - 1, 1).getValues(); // D

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
