/**
 * ADMIN CRM — Office Молодость v6.0
 * Исправленная версия на базе doc_edda4b78dc26_script.json
 */

const ADMIN_SPREADSHEET_ID = "1AUCWikqIhAGxXKVvTm1SLJjmLoUh0Y_c7Sry8X8cXos";
const DATA_ROWS = 1000;

const EMPLOYEES = [
  {name: "Тёмыч",  id: "1VCVAZhTl4cv9T1J4AyzknYqekMy6ZEymlQRkvOB4yJE"},
  {name: "Влад",   id: "1Lt9BmIVShNFserfYacxII6WjGPDn5ObNeiOo9z63oI8"},
  {name: "Соня",   id: "1U2uq6xhVXxcUvN3bTwj0eleQZ5mUm4c4OCpwN-deqoA"},
  {name: "Костян", id: "1CcaWPBvdPZ5WwegxDKOQhxljvMHwPZnYLgpl05QVB1c"},
  {name: "Денишк", id: "1pYRyigxMNSmrqr92RZ9-I5rbehErgX-PSiRbhn4Z4qU"}
];

const ADMIN_THEME = {
  headerBg: "#1a1a1a",
  headerText: "#00ff88",
  evenBg: "#2d2d2d",
  oddBg: "#1a1a1a",
  bodyText: "#e0e0e0",
  border: "#3a3a3a"
};

const EMPLOYEE_HEADERS = ["Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"];

const ACTIVE_KEEP_STATUSES = ["📝 Заявка", "🎫 Ожидает билеты", "🚗 Ожидает выезда", "🚀 В пути", "🏛️ В военкомате", "🔍 На проверке", "✅ Подписан"];

function initSystem() {
  initAdminTable_()';
  SpreadsheetApp.getActive().toast("✅ Система готова!", "Готово", 5);
}

function initAdminTable() {
  initAdminTable_();
}

function initAdminTable_() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);

  setupAllLeadsSheet_(ss);
  setupEmployeesSheet_(ss);
  setupActiveCandidatesSheet_(ss);
  setupStatsSheet_(ss);
  
  syncAllResponsesFromEmployees();
  syncActiveCandidates();
  updateAllStats();
  
  SpreadsheetApp.flush();
}

function setupAllLeadsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, "Все лиды");
  const headers = ["ID", "Сотрудник", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"];
  ensureHeaders_(sheet, headers, ADMIN_THEME);
  applyDarkTheme_(sheet, headers.length);
}

function setupEmployeesSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, "Сотрудники");
  const headers = ["Сотрудник", "ID таблицы", "Ссылка", "Статус", "Дата добавления"];
  ensureHeaders_(sheet, headers, ADMIN_THEME);
  
  sheet.getRange(1, 1, 1, headers.length).createFilter();
  
  for (let i = 0; i < EMPLOYEES.length; i++) {
    const emp = EMPLOYEES[i];
    const row = i + 2;
    sheet.getRange(row, 1, 1, headers.length).setValues([[emp.name, emp.id, "=HYPERLINK(\"https://docs.google.com/spreadsheets/d/"+emp.id+"\",\"Ссылка\")", "✅ Активен", "=TODAY()"]]);
  }
  
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 150);
}

function setupActiveCandidatesSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, "Активные лиды");
  const headers = ["ID", "Сотрудник", "Дата", "ФИО", "Телефон", "Статус", "Заметки"];
  ensureHeaders_(sheet, headers, ADMIN_THEME);
  sheet.getRange(1, 1, 1, headers.length).createFilter();
}

function setupStatsSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, "Статистика");
  const headers = ["Период", "Сотрудник", "Всего", "Подписано", "Отказ", "Связь", "Активные", "Успех %", "Дата"];
  ensureHeaders_(sheet, headers, ADMIN_THEME);
}

function syncActiveCandidates() {
  const admin = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet = admin.getSheetByName("Активные лиды");
  if (!sheet) return;

  const activeSet = new Set(ACTIVE_KEEP_STATUSES.map(s => s.toLowerCase().trim()));
  
  let data = [];
  let idCounter = 1;
  EMPLOYEES.forEach(function(emp) {
    try {
      const ss = SpreadsheetApp.openById(emp.id);
      const leadSheet = ss.getSheetByName("Лиды") || ss.getSheets()[0];
      if (!leadSheet) return;
      
      const lr = safeLastRow_(leadSheet);
      if (lr <= 1) return;
      
      const rows = leadSheet.getRange(2, 1, lr - 1, EMPLOYEE_HEADERS.length).getValues();
      rows.forEach(function(r) {
        const status = String(r[6] || "").trim();
        if (activeSet.has(status.toLowerCase())) {
          data.push([idCounter++, emp.name, r[0]||"", r[3]||"", r[4]||"", r[6]||"", r[7]||""]);
        }
      });
    } catch(e) { Logger.log("Ошибка "+emp.name+": "+e); }
  });
  
  sheet.clearContents();
  sheet.appendRow(["ID", "Сотрудник", "Дата", "ФИО", "Телефон", "Статус", "Заметки"]);
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 7).setValues(data);
  }
  applyDarkTheme_(sheet, 7);
}

function syncAllResponsesFromEmployees() {
  const admin = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(admin, "Все лиды");
  
  let data = [];
  let idCounter = 1;
  EMPLOYEES.forEach(function(emp) {
    try {
      const ss = SpreadsheetApp.openById(emp.id);
      const leadSheet = ss.getSheetByName("Лиды") || ss.getSheets()[0];
      if (!leadSheet) return;
      
      const lr = safeLastRow_(leadSheet);
      if (lr <= 1) return;
      
      const rows = leadSheet.getRange(2, 1, lr - 1, EMPLOYEE_HEADERS.length).getValues();
      rows.forEach(function(r) {
        if (!r[3] && !r[4]) return;
        data.push([idCounter++, emp.name, r[0]||"", r[1]||"", r[2]||"", r[3]||"", r[4]||"", r[5]||"", r[6]||"", r[7]||""]);
      });
    } catch(e) { Logger.log("Ошибка "+emp.name+": "+e); }
  });
  
  sheet.clearContents();
  sheet.appendRow(["ID", "Сотрудник", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"]);
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 10).setValues(data);
  }
  applyDarkTheme_(sheet, 10);
}

function updateAllStats() {
  const admin = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet = getOrCreateSheet_(admin, "Статистика");
  
  let data = [];
  EMPLOYEES.forEach(function(emp) {
    try {
      const ss = SpreadsheetApp.openById(emp.id);
      const leadSheet = ss.getSheetByName("Лиды") || ss.getSheets()[0];
      if (!leadSheet) return;
      
      const lr = safeLastRow_(leadSheet);
      if (lr <= 1) return;
      
      const rows = leadSheet.getRange(2, 1, lr - 1, EMPLOYEE_HEADERS.length).getValues();
      let total = 0, connected = 0, refused = 0, active = 0, connectionCount = 0;
      
      rows.forEach(function(r) {
        if (!r[3] && !r[4]) return;
        total++;
        const s = String(r[6] || "").toLowerCase().trim();
        if (s.includes("✅") || s.includes("подпис")) connected++;
        if (s.includes("🔴") || s.includes("❌") || s.includes("нд")) refused++;
        if (s.includes("💬") || s.includes("🟡")) connectionCount++;
        if (ACTIVE_KEEP_STATUSES.some(status => s.includes(status.toLowerCase()))) active++;
      });
      
      data.push(["Месяц", emp.name, total, connected, refused, connectionCount, active, total>0 ? Math.round(100*connected/total) : 0, "=TODAY()"]);
    } catch(e) { Logger.log("Ошибка "+emp.name+": "+e); }
  });
  
  sheet.clearContents();
  sheet.appendRow(["Период", "Сотрудник", "Всего", "Подписано", "Отказ", "Связь", "Активные", "Успех %", "Дата"]);
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 9).setValues(data);
  }
  applyDarkTheme_(sheet, 9);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Админ: Офис Молодость")
    .addItem("Создать структуру", "initSystem")
    .addItem("Обновить все данные", "syncAllResponsesFromEmployees")
    .addItem("Обновить активных", "syncActiveCandidates")
    .addToUi();
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaders_(sheet, headers, theme) {
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  let same = true;
  for (let i = 0; i < headers.length; i++) {
    if (current[i] !== headers[i]) {
      same = false;
      break;
    }
  }
  if (!same) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  setHeaders_(sheet, headers, theme);
}

function setHeaders_(sheet, headers, theme) {
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight("bold").setFontSize(11).setFontFamily("Segoe UI")
    .setBackground(theme.headerBg).setFontColor(theme.headerText)
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  sheet.setFrozenRows(1);
}

function applyDarkTheme_(sheet, cols) {
  const rowCount = DATA_ROWS;
  const bodyRows = rowCount - 1;

  sheet.getRange(2, 1, bodyRows, cols).setFontColor(ADMIN_THEME.bodyText)
    .setFontFamily("Segoe UI").setFontSize(10);

  for (let row = 2; row <= rowCount; row++) {
    sheet.getRange(row, 1, 1, cols).setBackground(row % 2 === 0 ? ADMIN_THEME.evenBg : ADMIN_THEME.oddBg);
  }

  sheet.getRange(1, 1, rowCount, cols).setBorder(true, true, true, true, true, true);
}

function safeLastRow_(sheet) {
  if (!sheet || typeof sheet.getLastRow !== "function") return 1;
  return sheet.getLastRow() || 1;
}
