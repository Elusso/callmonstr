// ============================================================
// CALLMONSTR ADMIN v8.0 — ИСПРАВЛЕННАЯ ВЕРСИЯ ДЛЯ ОФИСА МОЛОДОСТЬ
// ============================================================
const ADMIN_SPREADSHEET_ID = "1AUCWikqIhAGxXKVvTm1SLJjmLoUh0Y_c7Sry8X8cXos";
const DATA_ROWS = 1000;

const EMPLOYEES = [
  { name: "Тёмыч",  id: "1VCVAZhTl4cv9T1J4AyzknYqekMy6ZEymlQRkvOB4yJE" },
  { name: "Влад",   id: "1Lt9BmIVShNFserfYacxII6WjGPDn5ObNeiOo9z63oI8" },
  { name: "Соня",   id: "1U2uq6xhVXxcUvN3bTwj0eleQZ5mUm4c4OCpwN-deqoA" },
  { name: "Костян", id: "1CcaWPBvdPZ5WwegxDKOQhxljvMHwPZnYLgpl05QVB1c" },
  { name: "Денчик", id: "1pYRyigxMNSmrqr92RZ9-I5rbehErgX-PSiRbhn4Z4qU" }
];

const EMPLOYEE_HEADERS = ["Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"];
const ALL_LEADS_HEADERS = ["Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Сотрудник", "Заметки"];
const ACTIVE_HEADERS = ["Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Статус", "Заметки"];
const KOMISS_HEADERS = ["Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Статус", "Заметки", "Причина"];
const STATS_HEADERS = ["Период", "Сотрудник", "Всего", "Дозвон", "НД", "Новый", "ДУМ", "Отказ", "Слив", "Связь", "В пути", "Подписан", "Комиссован", "% дозвона", "Дата"];

// Активные — только те кто в процессе. Без Отказов и Сливов.
const ACTIVE_STATUSES = [
  "📝 Заявка", "🎫 Ожидает билеты", "🚗 Ожидает выезда",
  "🚀 В пути", "🏛 В военкомате", "🔍 На проверке",
  "✅ ПОДПИСАН", "✅ Подписан"
];

const VALID_STATUSES = [
  "⚪️ Новый", "🔴 НД", "🤔 ДУМ", "🟡 ПЕРЕЗВОНИТЬ", "💬 СВЯЗЬ МЕССЕНДЖЕР",
  "✅ ПОДПИСАН", "✅ Подписан", "⚫️ ОТКАЗ", "❌ Отказ", "💧 Слив",
  "📝 Заявка", "🎫 Ожидает билеты", "🚗 Ожидает выезда", "🚀 В пути",
  "🏛 В военкомате", "🔍 На проверке", "🎗 Комиссован"
];

// ============================================================
// ИНИЦИАЛИЗАЦИЯ
// ============================================================
function initFullSystem() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  setupSheets_(ss);
  setupTriggers_();
  syncAllLeads_();
  syncActiveCandidates_();
  syncKomissovannye_();
  updateStats_();
  ss.toast("✅ Система готова!", "Готово", 5);
}

function setupSheets_(ss) {
  ensureSheet_(ss, "Все лиды", ALL_LEADS_HEADERS);
  ensureSheet_(ss, "👥 Сотрудники", ["Сотрудник", "ID таблицы", "Ссылка", "Статус", "Дата"]);
  ensureSheet_(ss, "📊 Статистика", STATS_HEADERS);
  ensureSheet_(ss, "🎯 Активные кандидаты", ACTIVE_HEADERS);
  ensureSheet_(ss, "🎗 Комиссованные", KOMISS_HEADERS);

  const empSheet = ss.getSheetByName("👥 Сотрудники");
  EMPLOYEES.forEach(function(emp, i) {
    const row = i + 2;
    if (!empSheet.getRange(row, 1).getValue()) {
      empSheet.getRange(row, 1, 1, 5).setValues([[emp.name, emp.id,
        '=HYPERLINK("https://docs.google.com/spreadsheets/d/' + emp.id + '")', "✅ Активен", new Date()]]);
    }
  });
}

// ============================================================
// БЕЗОПАСНАЯ ЗАПИСЬ
// ============================================================
function safeSetValues_(sheet, range, values) {
  try { range.setValues(values); } catch (e) {
    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        try { sheet.getRange(range.getRow() + i, range.getColumn() + j).setValue(values[i][j]); } catch (e2) {}
      }
    }
  }
}

function normalizeStatus_(status) {
  if (!status) return "⚪️ Новый";
  const s = String(status).trim();
  if (VALID_STATUSES.indexOf(s) >= 0) return s;
  const lower = s.toLowerCase();
  for (const vs of VALID_STATUSES) {
    if (vs.toLowerCase().includes(lower) || lower.includes(vs.toLowerCase())) return vs;
  }
  return "⚪️ Новый";
}

// ============================================================
// СИНХРОНИЗАЦИЯ ВСЕХ ЛИДОВ
// ============================================================
function syncAllLeads_() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Все лиды");
  if (!sheet) return;

  sheet.getRange("G2:G" + DATA_ROWS).clearDataValidations();
const existing = {};
  const lastRow = safeLastRow_(sheet);
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 9).getValues().forEach(function(r, i) {
      const key = normalizeKey_(r[3], r[4]);
      if (key) existing[key] = i + 2;
    });
  }

  const updates = [];
  const appends = [];

  EMPLOYEES.forEach(function(emp) {
    try {
      const ess = SpreadsheetApp.openById(emp.id);
      const leadSheet = ess.getSheetByName("Лиды") || ess.getSheets()[0];
      if (!leadSheet) return;
      const lr = safeLastRow_(leadSheet);
      if (lr <= 1) return;

      leadSheet.getRange(2, 1, lr - 1, EMPLOYEE_HEADERS.length).getValues().forEach(function(r) {
        const key = normalizeKey_(r[3], r[4]);
        if (!key) return;
        const status = normalizeStatus_(r[6]);
        const dateFromEmployee = r[0] ? new Date(r[0]) : null;
        const date = (dateFromEmployee && !isNaN(dateFromEmployee.getTime())) ? dateFromEmployee : new Date();
        const row = [date, r[1] || "", r[2] || "", r[3] || "", r[4] || "", r[5] || "", status, emp.name, r[7] || ""];
        if (existing[key]) updates.push({ ri: existing[key], row: row });
        else appends.push(row);
      });
    } catch (e) { Logger.log("Ошибка синхронизации " + emp.name + ": " + e.message); }
  });

  updates.forEach(function(u) { safeSetValues_(sheet, sheet.getRange(u.ri, 1, 1, 9), [u.row]); });
  if (appends.length > 0) {
    safeSetValues_(sheet, sheet.getRange(safeLastRow_(sheet) + 1, 1, appends.length, 9), appends);
  }

  const rule = SpreadsheetApp.newDataValidation().requireValueInList(VALID_STATUSES, true).setAllowInvalid(false).build();
  sheet.getRange("G2:G" + DATA_ROWS).setDataValidation(rule);
  applyStyle_(sheet, ALL_LEADS_HEADERS.length);
}

// ============================================================
// СИНХРОНИЗАЦИЯ АКТИВНЫХ КАНДИДАТОВ (БЕЗ ОТКАЗОВ И СЛИВОВ)
// ============================================================
function syncActiveCandidates_() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("🎯 Активные кандидаты");
  if (!sheet) return;

  sheet.getRange("G2:G" + DATA_ROWS).clearDataValidations();

  const existingMap = {};
  const lastRow = safeLastRow_(sheet);
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, ACTIVE_HEADERS.length).getValues().forEach(function(r) {
      const key = normalizeKey_(r[2], r[3]);
      if (key) existingMap[key] = { row: r, rowIndex: 0 };
    });
  }

  const activeSet = new Set(ACTIVE_STATUSES.map(s => s.toLowerCase()));
  const updates = [];
  const appends = [];

  EMPLOYEES.forEach(function(emp) {
    try {
      const ess = SpreadsheetApp.openById(emp.id);
      const leadSheet = ess.getSheetByName("Лиды") || ess.getSheets()[0];
      if (!leadSheet) return;
      const lr = safeLastRow_(leadSheet);
      if (lr <= 1) return;

      leadSheet.getRange(2, 1, lr - 1, EMPLOYEE_HEADERS.length).getValues().forEach(function(r) {
        const fio = r[3];
        const phone = r[4];
        const status = normalizeStatus_(r[6]);
        if (!activeSet.has(status.toLowerCase())) return;
        const key = normalizeKey_(fio, phone);
        if (!key) return;

        const dateFromEmployee = r[0] ? new Date(r[0]) : null;
        const date = (dateFromEmployee && !isNaN(dateFromEmployee.getTime())) ? dateFromEmployee : new Date();
        const old = existingMap[key];
        const payload = [old ? old.row[0] : date, emp.name, fio, phone, r[1] || "", r[2] || "", status, r[7] || ""];

        if (old) updates.push({ key: key, row: payload });
        else appends.push(payload);
      });
    } catch (e) { Logger.log("Ошибка активных " + emp.name + ": " + e.message); }
  });

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, ACTIVE_HEADERS.length).getValues().forEach(function(r, i) {
      const key = normalizeKey_(r[2], r[3]);
      if (key && existingMap[key]) existingMap[key].rowIndex = i + 2;
    });
  }
updates.forEach(function(u) {
    const ri = existingMap[u.key] ? existingMap[u.key].rowIndex : null;
    if (ri) safeSetValues_(sheet, sheet.getRange(ri, 1, 1, ACTIVE_HEADERS.length), [u.row]);
  });

  if (appends.length > 0) {
    safeSetValues_(sheet, sheet.getRange(safeLastRow_(sheet) + 1, 1, appends.length, ACTIVE_HEADERS.length), appends);
  }

  const rule = SpreadsheetApp.newDataValidation().requireValueInList(ACTIVE_STATUSES, true).setAllowInvalid(false).build();
  sheet.getRange("G2:G" + DATA_ROWS).setDataValidation(rule);
  applyStyle_(sheet, ACTIVE_HEADERS.length);
}

// ============================================================
// СИНХРОНИЗАЦИЯ КОМИССОВАННЫХ
// ============================================================
function syncKomissovannye_() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("🎗 Комиссованные");
  if (!sheet) return;

  sheet.getRange("G2:G" + DATA_ROWS).clearDataValidations();

  const existingMap = {};
  const lastRow = safeLastRow_(sheet);
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, KOMISS_HEADERS.length).getValues().forEach(function(r) {
      const key = normalizeKey_(r[2], r[3]);
      if (key) existingMap[key] = { row: r, rowIndex: 0 };
    });
  }

  const updates = [];
  const appends = [];

  EMPLOYEES.forEach(function(emp) {
    try {
      const ess = SpreadsheetApp.openById(emp.id);
      const leadSheet = ess.getSheetByName("Лиды") || ess.getSheets()[0];
      if (!leadSheet) return;
      const lr = safeLastRow_(leadSheet);
      if (lr <= 1) return;

      leadSheet.getRange(2, 1, lr - 1, EMPLOYEE_HEADERS.length).getValues().forEach(function(r) {
        const fio = r[3];
        const phone = r[4];
        const status = normalizeStatus_(r[6]);
        if (status !== "🎗 Комиссован") return;
        const key = normalizeKey_(fio, phone);
        if (!key) return;

        const dateFromEmployee = r[0] ? new Date(r[0]) : null;
        const date = (dateFromEmployee && !isNaN(dateFromEmployee.getTime())) ? dateFromEmployee : new Date();
        const old = existingMap[key];
        const payload = [old ? old.row[0] : date, emp.name, fio, phone, r[1] || "", r[2] || "", status, r[7] || "", "Комиссия"];

        if (old) updates.push({ key: key, row: payload });
        else appends.push(payload);
      });
    } catch (e) { Logger.log("Ошибка комиссованных " + emp.name + ": " + e.message); }
  });

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, KOMISS_HEADERS.length).getValues().forEach(function(r, i) {
      const key = normalizeKey_(r[2], r[3]);
      if (key && existingMap[key]) existingMap[key].rowIndex = i + 2;
    });
  }

  updates.forEach(function(u) {
    const ri = existingMap[u.key] ? existingMap[u.key].rowIndex : null;
    if (ri) safeSetValues_(sheet, sheet.getRange(ri, 1, 1, KOMISS_HEADERS.length), [u.row]);
  });

  if (appends.length > 0) {
    safeSetValues_(sheet, sheet.getRange(safeLastRow_(sheet) + 1, 1, appends.length, KOMISS_HEADERS.length), appends);
  }

  applyStyle_(sheet, KOMISS_HEADERS.length);
}

// ============================================================
// СТАТИСТИКА
// ============================================================
function updateStats_() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("📊 Статистика");
  if (!sheet) return;

  const rows = [];
  EMPLOYEES.forEach(function(emp) {
    try {
      const ess = SpreadsheetApp.openById(emp.id);
      const leadSheet = ess.getSheetByName("Лиды") || ess.getSheets()[0];
      const allLeads = [];
      if (leadSheet) {
        const lr = safeLastRow_(leadSheet);
        if (lr > 1) {
          leadSheet.getRange(2, 1, lr - 1, EMPLOYEE_HEADERS.length).getValues().forEach(function(r) {
            if (!r[3] && !r[4]) return;
            allLeads.push({ date: parseDate_(r[0]), status: String(r[6] || "").toLowerCase().trim() });
          });
        }
      }
[{ key: "День", filter: isToday_ }, { key: "Неделя", filter: function(d) { return isInLastNDays_(d, 7); } }, { key: "Месяц", filter: function(d) { return isInLastNDays_(d, 30); } }].forEach(function(p) {
        const filtered = allLeads.filter(function(x) { return p.filter(x.date); });
        const s = computeStats_(filtered);
        rows.push([p.key, emp.name, s.total, s.dozvon, s.nd, s.noviy, s.dum, s.otkaz, s.sliv, s.svyaz, s.vputi, s.podpisan, s.komissovan, s.pct, new Date()]);
      });
    } catch (e) { Logger.log("Ошибка статистики " + emp.name + ": " + e.message); }
  });

  const old = safeLastRow_(sheet);
  if (old > 1) sheet.getRange(2, 1, old - 1, STATS_HEADERS.length).clearContent();
  if (rows.length > 0) safeSetValues_(sheet, sheet.getRange(2, 1, rows.length, STATS_HEADERS.length), rows);
  sheet.getRange("N2:N" + DATA_ROWS).setNumberFormat("0.00");
  applyStyle_(sheet, STATS_HEADERS.length);
}

function computeStats_(leads) {
  const r = { total: 0, dozvon: 0, nd: 0, noviy: 0, dum: 0, otkaz: 0, sliv: 0, svyaz: 0, vputi: 0, podpisan: 0, komissovan: 0, pct: 0 };
  leads.forEach(function(item) {
    const s = item.status;
    r.total++;
    if (/нд/i.test(s)) r.nd++;
    if (/нов/i.test(s)) r.noviy++;
    if (/дум/i.test(s)) r.dum++;
    if (/слив/i.test(s)) r.sliv++;
    else if (/отказ/i.test(s)) r.otkaz++;
    if (/связь/i.test(s)) r.svyaz++;
    if (/в пути|ожидает выезда/i.test(s)) r.vputi++;
    if (/подписан/i.test(s)) r.podpisan++;
    if (/комиссован/i.test(s)) r.komissovan++;
  });
  r.dozvon = Math.max(r.total - r.nd, 0);
  r.pct = r.total > 0 ? Math.round((r.dozvon / r.total) * 100 * 100) / 100 : 0;
  return r;
}

// ============================================================
// ТРИГГЕРЫ И МЕНЮ
// ============================================================
function setupTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() !== "backupAllTables_") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("syncActiveCandidates_").timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger("syncAll_").timeBased().everyHours(1).create();
  ScriptApp.newTrigger("onOpen_").forSpreadsheet(ADMIN_SPREADSHEET_ID).onOpen().create();
}

function onOpen(e) { if (e && e.source && e.source.getId() === ADMIN_SPREADSHEET_ID) { addMenu_(); syncActiveCandidates_(); } }
function onOpen_() { addMenu_(); syncActiveCandidates_(); }
function syncAll_() { syncAllLeads_(); syncActiveCandidates_(); syncKomissovannye_(); updateStats_(); }

function addMenu_() {
  try {
    SpreadsheetApp.getUi().createMenu("🚀 Админ")
      .addItem("🔄 Обновить всё", "syncAll_")
      .addItem("📊 Статистика", "updateStats_")
      .addItem("🎯 Активные кандидаты", "syncActiveCandidates_")
      .addItem("🎗 Комиссованные", "syncKomissovannye_")
      .addToUi();
  } catch (e) {}
}

// ============================================================
// ВСПОМОГАТЕЛЬНЫЕ
// ============================================================
function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  let same = true;
  for (let i = 0; i < headers.length; i++) if (current[i] !== headers[i]) { same = false; break; }
  if (!same) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold").setFontSize(11).setFontFamily("Comfortaa")
    .setBackground("#16213e").setFontColor("#ffffff")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  sheet.setFrozenRows(1);
}
function applyStyle_(sheet, cols) {
  const lastRow = safeLastRow_(sheet);
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, cols).setFontFamily("Comfortaa").setFontSize(10);
  for (let r = 2; r <= Math.max(lastRow, 1); r++) {
    try { sheet.getRange(r, 1, 1, cols).setBackground(r % 2 === 0 ? "#1a1a2e" : "#16213e"); } catch (e) {}
  }
  sheet.getRange(1, 1, Math.max(lastRow, DATA_ROWS), cols).setBorder(true, true, true, true, true, true);
}

function safeLastRow_(sheet) { return sheet && sheet.getLastRow ? sheet.getLastRow() || 1 : 1; }
function normalizeKey_(fio, phone) { const f = String(fio || "").trim().toLowerCase(); const p = String(phone || "").replace(/\D/g, ""); return (!f && !p) ? "" : f + "|" + p; }
function parseDate_(v) { if (!v) return null; if (v instanceof Date && !isNaN(v)) return v; const d = new Date(v); return isNaN(d.getTime()) ? null : d; }
function isToday_(d) { if (!d) return false; const n = new Date(); return d.getFullYear() === n.getFullYear() && d.getMonth() === n.getMonth() && d.getDate() === n.getDate(); }
function isInLastNDays_(d, days) { if (!d) return false; const n = new Date(); const start = new Date(n.getFullYear(), n.getMonth(), n.getDate() - days + 1); return d >= start && d <= n; }

function backupAllTables_() {
  const folder = DriveApp.getFoldersByName("CallMonstr Backups").hasNext() ?
    DriveApp.getFoldersByName("CallMonstr Backups").next() : DriveApp.createFolder("CallMonstr Backups");
  const d = Utilities.formatDate(new Date(), "GMT+3", "dd.MM.yyyy");
  EMPLOYEES.forEach(function(emp) { try { DriveApp.getFileById(emp.id).makeCopy("📦 " + emp.name + " " + d).moveTo(folder); } catch (e) {} });
  try { DriveApp.getFileById(ADMIN_SPREADSHEET_ID).makeCopy("📦 АДМИН " + d).moveTo(folder); } catch (e) {}
}
