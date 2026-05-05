/**
 * ADMIN CRM — Офис Молодость (v2.2)
 * =================================
 * Главная CRM для владельца офиса.
 * Считывает лиды со всех сотрудников по фиксированной структуре:
 *   A=Дата, B=Вакансия, C=Город, D=ФИО, E=Телефон, F=Возраст, G=Статус, H=Заметки
 * 
 * Листы:
 * - Все лиды: все лиды, отсортированные по дате (новее сверху), красивое оформление
 * - Статистика: сводные показатели по всему офису
 * - Активные лиды: только "в активной работе" (в пути, ожидает выезда и т.д.)
 * 
 * ВАЖНО: ТИЕРАВ НЕТ! Это личное у Влада, не считываем!
 */

// Global Config
const ADMIN_SPREADSHEET_ID = "1AUCWikqIhAGxXKVvTm1SLJjmLoUh0Y_c7Sry8X8cXos";

const THEME = {
  bgDeep: "#1a1a1a",
  bgMid: "#2d2d2d",
  accent: "#00ff88",
  text: "#e0e0e0",
  font: "Comfortaa"
};

// Сотрудники
const EMPLOYEES = [
  {name: "Тёмыч",  id: "1VCVAZhTl4cv9T1J4AyzknYqekMy6ZEymlQRkvOB4yJE"},
  {name: "Влад",   id: "1Lt9BmIVShNFserfYacxII6WjGPDn5ObNeiOo9z63oI8"},
  {name: "Соня",   id: "1U2uq6xhVXxcUvN3bTwj0eleQZ5mUm4c4OCpwN-deqoA"},
  {name: "Костян", id: "1CcaWPBvdPZ5WwegxDKOQhxljvMHwPZnYLgpl05QVB1c"},
  {name: "Денишк", id: "1pYRyigxMNSmrqr92RZ9-I5rbehErgX-PSiRbhn4Z4qU"}
];

// Константы листов
const LEADS_SHEET = "Все лиды";
const STATS_SHEET = "Статистика";
const ACTIVES_SHEET = "Активные лиды";

// Индексы колонок (0-based)
const COLS = { date: 0, vacancy: 1, city: 2, fio: 3, phone: 4, age: 5, status: 6, notes: 7 };

// Заголовки
function LEADS_HEADER() { return ["ID", "Сотрудник", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"]; }
function STATS_HEADER() { return ["Дата", "Всего", "Подписано", "Отказ", "Свяьь", "Активные", "Успех %"]; }
function ACTIVES_HEADER() { return ["ID", "Сотрудник", "Дата", "ФИО", "Телефон", "Статус", "Заметки"]; }

// Статусы "в активной работе"
const ACTIVE_STATUSES = ["🚀 В пути", "🚗 Ожидает выезда", "🏛️ В военкомате", "🔍 На проверке"];

// Стилизация (темная тема, чередующаяся заливка)
function styleHeader(range) {
  range.setBackground("#1a1a1a")
       .setFontColor("#00ff88")
       .setFontFamily("Comfortaa")
       .setFontWeight("bold")
       .setFontSize(12);
}

function styleCell(row, col, isHeader = false) {
  const cell = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(LEADS_SHEET)
    .getRange(row, col);
  
  if (isHeader) {
    cell.setBackground("#1a1a1a")
        .setFontColor("#00ff88")
        .setFontWeight("bold");
  } else {
    // чередующаяся заливка
    const mod = (row - 2) % 2;
    if (mod === 0) {
      cell.setBackground("#2d2d2d");
    } else {
      cell.setBackground("#1a1a1a");
    }
    cell.setFontColor("#e0e0e0");
  }
}

function setColumnsWidth(sheet, widths) {
  widths.forEach((w, i) => { try { sheet.setColumnWidth(i + 1, w); } catch {} });
  if (sheet.getLastRow() >= 1) sheet.setRowHeight(1, 40);
}

// Получение данных из таблицы сотрудника
function readEmployeeSheet(emp) {
  try {
    const ss = SpreadsheetApp.openById(emp.id);
    const sheet = ss.getSheets()[0];
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (r.length < 8) continue;
      
      rows.push({
        name: emp.name,
        date: r[COLS.date] || "",
        vacancy: r[COLS.vacancy] || "",
        city: r[COLS.city] || "",
        fio: r[COLS.fio] || "",
        phone: r[COLS.phone] || "",
        age: r[COLS.age] || "",
        status: r[COLS.status] || "",
        notes: r[COLS.notes] || ""
      });
    }
    return rows;
  } catch (e) {
    Logger.log(`⚠️ Ошибка при чтении ${emp.name}: ${e}`);
    return [];
  }
}

// Сбор лидов со всех сотрудников
function getAllLeads() {
  let all = [];
  EMPLOYEES.forEach(emp => {
    const leads = readEmployeeSheet(emp);
    all = all.concat(leads);
  });
  Logger.log(`✅ Считано лидов: ${all.length}`);
  return all;
}

// Создание листов
function initSystem() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheets = ss.getSheets().map(s => s.getName());
  
  // "Все лиды"
  if (!sheets.includes(LEADS_SHEET)) {
    const sh = ss.insertSheet(LEADS_SHEET);
    sh.appendRow(LEADS_HEADER());
    styleHeader(sh.getRange(1, 1, 1, LEADS_HEADER().length));
    sh.setFrozenRows(1);
    setColumnsWidth(sh, [50, 120, 100, 180, 120, 220, 140, 70, 120, 180]);
    Logger.log("✅ Лист 'Все лиды' создан");
  }
  
  // "Статистика"
  if (!sheets.includes(STATS_SHEET)) {
    const sh = ss.insertSheet(STATS_SHEET);
    sh.appendRow(STATS_HEADER());
    styleHeader(sh.getRange(1, 1, 1, STATS_HEADER().length));
    sh.setFrozenRows(1);
    setColumnsWidth(sh, [120, 80, 90, 90, 90, 80, 80]);
    Logger.log("✅ Лист 'Статистика' создан");
  }
  
  // "Активные лиды"
  if (!sheets.includes(ACTIVES_SHEET)) {
    const sh = ss.insertSheet(ACTIVES_SHEET);
    sh.appendRow(ACTIVES_HEADER());
    styleHeader(sh.getRange(1, 1, 1, ACTIVES_HEADER().length));
    sh.setFrozenRows(1);
    setColumnsWidth(sh, [50, 120, 100, 220, 140, 120, 180]);
    Logger.log("✅ Лист 'Активные лиды' создан");
  }
  
  return "✅ Листы созданы: Все лиды, Статистика, Активные лиды";
}

// Синхронизация и сбор лидов
function syncAllSheets() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const allLeads = getAllLeads();
  
  // 1. "Все лиды" — отсортировать по дате (новее сверху)
  const leadsSheet = ss.getSheetByName(LEADS_SHEET);
  if (!leadsSheet) return "❌ Лист 'Все лиды' не найден";
  
  // очистка
  if (leadsSheet.getLastRow() > 1) leadsSheet.deleteRows(2, leadsSheet.getLastRow() - 1);
  
  // сортировка по дате (новее сверху)
  allLeads.sort((a, b) => {
    const dA = Date.parse(a.date) || 0;
    const dB = Date.parse(b.date) || 0;
    return dB - dA;
  });
  
  if (allLeads.length > 0) {
    const data = allLeads.map((l, i) => [
      i + 1,
      l.name,
      l.date,
      l.vacancy,
      l.city,
      l.fio,
      l.phone,
      l.age,
      l.status,
      l.notes
    ]);
    
    const range = leadsSheet.getRange(2, 1, data.length, data[0].length);
    range.setValues(data);
    
    // применяем стили
    const startRow = 2;
    const rowCount = data.length;
    const colCount = data[0].length;
    
    // стилизация заголовков
    const headerRange = leadsSheet.getRange(1, 1, 1, colCount);
    headerRange.setBackground("#1a1a1a");
    headerRange.setFontColor("#00ff88");
    headerRange.setFontWeight("bold");
    
    // чередующаяся заливка
    for (let r = 0; r < rowCount; r++) {
      for (let c = 0; c < colCount; c++) {
        const cell = leadsSheet.getRange(startRow + r, c + 1);
        const isEven = (r % 2 === 0);
        cell.setBackground(isEven ? "#2d2d2d" : "#1a1a1a");
        cell.setFontColor("#e0e0e0");
      }
    }
    
    // границы для всех ячеек
    range.setBorder(true, true, true, true, true, true);
  }
  
  // 2. "Статистика"
  const statsSheet = ss.getSheetByName(STATS_SHEET);
  if (statsSheet) {
    statsSheet.clear();
    statsSheet.appendRow(STATS_HEADER());
    styleHeader(statsSheet.getRange(1, 1, 1, STATS_HEADER().length));
    
    // Подсчёт статистики
    const today = new Date();
    let total = allLeads.length;
    let connected = 0, refused = 0, contact = 0, activesCount = 0;
    
    for (let i = 0; i < allLeads.length; i++) {
      const st = allLeads[i].status;
      if (st.includes("✅")) connected++;
      if (st.includes("🔴") || st.includes("⚫") || st.includes("❌")) refused++;
      if (st.includes("💬") || st.includes("🟡")) contact++;
      if (ACTIVE_STATUSES.indexOf(st) >= 0) activesCount++;
    }
    
    const successRate = total > 0 ? Math.round(100 * connected / total) : 0;
    
    statsSheet.appendRow([
      today,
      total,
      connected,
      refused,
      contact,
      activesCount,
      successRate + "%"
    ]);
    
    // границы
    statsSheet.getRange(1, 1, 2, 7).setBorder(true, true, true, true, true, true);
  }
  
  // 3. "Активные лиды" — только "в активной работе"
  const activesSheet = ss.getSheetByName(ACTIVES_SHEET);
  if (activesSheet) {
    if (activesSheet.getLastRow() > 1) activesSheet.deleteRows(2, activesSheet.getLastRow() - 1);
    
    // Фильтр: только активные
    const actives = allLeads.filter(l => ACTIVE_STATUSES.indexOf(l.status) >= 0);
    
    if (actives.length > 0) {
      const data = actives.map((l, i) => [
        i + 1,
        l.name,
        l.date,
        l.fio,
        l.phone,
        l.status,
        l.notes
      ]);
      
      const range = activesSheet.getRange(2, 1, data.length, data[0].length);
      range.setValues(data);
      
      // стили
      const headerRange = activesSheet.getRange(1, 1, 1, data[0].length);
      headerRange.setBackground("#1a1a1a");
      headerRange.setFontColor("#00ff88");
      headerRange.setFontWeight("bold");
      
      // чередующаяся заливка
      for (let r = 0; r < data.length; r++) {
        for (let c = 0; c < data[0].length; c++) {
          const cell = activesSheet.getRange(2 + r, c + 1);
          const isEven = (r % 2 === 0);
          cell.setBackground(isEven ? "#2d2d2d" : "#1a1a1a");
          cell.setFontColor("#e0e0e0");
        }
      }
      
      range.setBorder(true, true, true, true, true, true);
    }
  }
  
  return "✅ Синхронизация: " + allLeads.length + " лидов";
}

// Выгрузка JSON
function exportJSON() {
  const allLeads = getAllLeads();
  const json = JSON.stringify(allLeads, null, 2);
  const file = DriveApp.createFile("admin_crm_leads.json", json, "application/json");
  return "✅ JSON выгружен: " + file.getUrl();
}

// Меню
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("_Офис Молодость Admin v2.2")
    .addItem("Создать листы", "initSystem")
    .addItem("Синхронизировать все таблицы", "syncAllSheets")
    .addItem("Экспорт JSON", "exportJSON")
    .addToUi();
}
