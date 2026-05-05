/**
 * ADMIN DASHBOARD — VLAD (CallMonstr v2.0)
 * ========================================
 * Google Apps Script для работы со всей CRM: админ + влав
 * Функции:
 * - initSystem()      — создание всех листов
 * - syncAllSheets()   — синхронизация всех таблиц
 * - generateReport()  — статистика по сотрудникам
 * - exportJSON()      — выгрузка JSON для внешних систем
 * - onOpen()          — меню в Google Sheets
 */

const ADMIN_SHEET_ID = "1admin1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"; // замени на свой ID
const VLAD_SHEET_ID  = "1vlad1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ";  // замени на свой ID

// Константы
const ADMIN_SHEET = "Все лиды";
const TIER_SHEET  = "Tier-list";
const WORKED_SHEET = "Отработать";
const STATS_SHEET = "Статистика";

// Статусы → тиер
const STATUS_TO_TIER = {
  "✅ Подписан": "SS",
  "✅ Подписаны": "SS",
  "✅ ПОДПИСАН": "SS",
  "🎗️ Комиссован": "SS",
  "🔴 НД": "S",
  "🤔 ДУМ": "S",
  "⚫ ОТКАЗ": "S",
  "❌ Отказ": "S",
  "🟡 ПЕРЕЗВОНИТЬ": "A",
  "💬 СВЯЗЬ МЕССЕНДЖЕР": "A",
  "🎫 Ожидает билеты": "A",
  "🚗 Ожидает выезда": "A",
  "🚀 В пути": "A",
  "🏛️ В военкомате": "A",
  "🔍 На проверке": "A",
  "📝 Заявка": "A",
  "default": "A"
};

// Заголовки колонок
function ADMIN_HEADERS() { return ["ID", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Тир", "Комментарий", "Дата напоминания"]; }
function TIER_HEADERS()  { return ["ID", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Тир", "Комментарий", "Дата напоминания"]; }
function WORKED_HEADERS() { return ["ID", "Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Статус", "Комментарий"]; }
function STATS_HEADERS() { return ["Дата", "Всего лидов", "T1 (SS)", "T2 (S)", "T3 (A)", "Подписано", "Отказ", "Связь", "Успешность"]; }

// Стилизация
function styleHeader(range) {
  range.setBackground("#1a1a1a")
       .setFontColor("#00ff88")
       .setFontFamily("Comfortaa")
       .setFontWeight("bold")
       .setFontSize(11);
}

function setColumnWidths(sheet, widths) {
  widths.forEach((w, i) => { try { sheet.setColumnWidth(i + 1, w); } catch {} });
  if (sheet.getLastRow() >= 1) sheet.setRowHeight(1, 40);
}

// Статус → тиер
function getStatusTier(status) {
  if (!status) return STATUS_TO_TIER["default"];
  return STATUS_TO_TIER[status.trim()] || STATUS_TO_TIER["default"];
}

// парсинг даты
function parseDate(val) {
  if (!val) return 0;
  const d = new Date(val);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

// Получение данных из таблицы
function readSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1);
}

// Синхронизация всех таблиц
function syncAllSheets() {
  const ss = SpreadsheetApp.openById(ADMIN_SHEET_ID);
  const sheet = ss.getSheetByName(ADMIN_SHEET);
  
  if (!sheet) return "❌ Таблица 'Все лиды' не найдена";
  
  const mainData = readSheet(ss, ADMIN_SHEET);
  if (mainData.length === 0) return "⏸ Нет данных в 'Все лиды'";
  
  // --- Tier-list ---
  let tierSheet = ss.getSheetByName(TIER_SHEET);
  if (!tierSheet) {
    tierSheet = ss.insertSheet(TIER_SHEET);
    tierSheet.appendRow(TIER_HEADERS());
    styleHeader(tierSheet.getRange(1, 1, 1, TIER_HEADERS().length));
    tierSheet.setFrozenRows(1);
    setColumnWidths(tierSheet, [60, 120, 180, 140, 220, 140, 80, 120, 60, 200, 150]);
  }
  
  // очистка
  if (tierSheet.getLastRow() > 1) tierSheet.deleteRows(2, tierSheet.getLastRow() - 1);
  
  const tierData = [];
  for (let i = 0; i < mainData.length; i++) {
    const row = mainData[i];
    if (row.length < 8) continue;
    
    const phone = row[5];
    if (!phone) continue;
    
    const status = row[7];
    const tier = getStatusTier(status);
    
    const tgt = [
      row[0] || `LID-${i+1}`,
      row[1] || '',
      row[2] || '',
      row[3] || '',
      row[4] || '',
      row[5] || '',
      row[6] || '',
      status,
      tier,
      row[8] || '',
      new Date()
    ];
    tierData.push(tgt);
  }
  
  if (tierData.length > 0) {
    tierSheet.getRange(2, 1, tierData.length, tierData[0].length).setValues(tierData);
    applyTierColors(tierSheet, 2, tierData.length);
  }
  
  // --- Сортировка ---
  sortTierList(ss);
  
  // --- Stats sheet ---
  let statsSheet = ss.getSheetByName(STATS_SHEET);
  if (!statsSheet) {
    statsSheet = ss.insertSheet(STATS_SHEET);
    statsSheet.appendRow(STATS_HEADERS());
    styleHeader(statsSheet.getRange(1, 1, 1, STATS_HEADERS().length));
    statsSheet.setFrozenRows(1);
    setColumnWidths(statsSheet, [120, 90, 80, 80, 80, 90, 90, 90, 80]);
  }
  
  // подсчёт статистики
  const stats = calculateStats(tierData);
  statsSheet.appendRow([
    new Date(),
    stats.total,
    stats.T1,
    stats.T2,
    stats.T3,
    stats.connected,
    stats.refused,
    stats.contact,
    stats.successRate + "%"
  ]);
  
  return "✅ Синхронизация выполнена: " + tierData.length + " записей";
}

// Сортировка по тиеру (SS → S → A)
function sortTierList(ss) {
  const tierSheet = ss.getSheetByName(TIER_SHEET);
  if (!tierSheet) return "❌ Tier-list не найден";
  
  const data = tierSheet.getDataRange().getValues();
  if (data.length <= 1) return "⏸ Нет данных";
  
  const rows = data.slice(1);
  
  rows.sort((a, b) => {
    const tierA = getTierRank(a[8]);
    const tierB = getTierRank(b[8]);
    if (tierA !== tierB) return tierA - tierB;
    const dateA = parseDate(a[1]);
    const dateB = parseDate(b[1]);
    return dateB - dateA;
  });
  
  if (tierSheet.getLastRow() > 1) tierSheet.deleteRows(2, tierSheet.getLastRow() - 1);
  if (rows.length > 0) {
    tierSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    applyTierColors(tierSheet, 2, rows.length);
  }
  
  return "⬆️ Сортировка: " + rows.length + " записей";
}

// Расчёт статистики
function calculateStats(data) {
  let total = data.length;
  let T1 = 0, T2 = 0, T3 = 0;
  let connected = 0, refused = 0, contact = 0;
  
  for (let i = 0; i < data.length; i++) {
    const status = data[i][7] || '';
    const tier = data[i][8] || '';
    
    if (tier === 'SS') T1++;
    else if (tier === 'S') T2++;
    else if (tier === 'A') T3++;
    
    if (status.includes('✅')) connected++;
    if (status.includes('🔴') || status.includes('⚫') || status.includes('❌') || status.includes('🤔')) refused++;
    if (status.includes('💬') || status.includes('🟡')) contact++;
  }
  
  const successRate = total > 0 ? Math.round(100 * connected / total) : 0;
  
  return {
    total: total,
    T1: T1,
    T2: T2,
    T3: T3,
    connected: connected,
    refused: refused,
    contact: contact,
    successRate: successRate
  };
}

// цвета тиеров
function applyTierColors(sheet, startRow, rowCount) {
  for (let r = 0; r < rowCount; r++) {
    const tier = sheet.getRange(startRow + r, 9).getValue();
    const cell = sheet.getRange(startRow + r, 9);
    
    if (tier === 'SS') cell.setBackground('#ff3f34');
    else if (tier === 'S') cell.setBackground('#ffa502');
    else if (tier === 'A') cell.setBackground('#ffff00');
    else cell.setBackground('#ffffff');
  }
}

function getTierRank(tier) {
  if (tier === 'SS') return 1;
  if (tier === 'S') return 2;
  if (tier === 'A') return 3;
  return 999;
}

// выгрузка JSON
function exportJSON() {
  const ss = SpreadsheetApp.openById(ADMIN_SHEET_ID);
  const sheet = ss.getSheetByName(TIER_SHEET);
  if (!sheet) return JSON.stringify({ error: "Tier-list not found" });
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return JSON.stringify({ error: "No data" });
  
  const json = [];
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    json.push(obj);
  }
  
  return JSON.stringify(json, null, 2);
}

// инициализация системы
function initSystem() {
  const ss = SpreadsheetApp.openById(ADMIN_SHEET_ID);
  
  // "Все лиды"
  let mainSheet = ss.getSheetByName(ADMIN_SHEET);
  if (!mainSheet) {
    mainSheet = ss.insertSheet(ADMIN_SHEET);
    mainSheet.appendRow(ADMIN_HEADERS());
    styleHeader(mainSheet.getRange(1, 1, 1, ADMIN_HEADERS().length));
    mainSheet.setFrozenRows(1);
    setColumnWidths(mainSheet, [60, 120, 180, 140, 220, 140, 80, 120, 60, 200, 150]);
  }
  
  // "Tier-list"
  if (!ss.getSheetByName(TIER_SHEET)) {
    const tSheet = ss.insertSheet(TIER_SHEET);
    tSheet.appendRow(TIER_HEADERS());
    styleHeader(tSheet.getRange(1, 1, 1, TIER_HEADERS().length));
    tSheet.setFrozenRows(1);
    setColumnWidths(tSheet, [60, 120, 180, 140, 220, 140, 80, 120, 60, 200, 150]);
  }
  
  // "Отработать"
  if (!ss.getSheetByName(WORKED_SHEET)) {
    const wSheet = ss.insertSheet(WORKED_SHEET);
    wSheet.appendRow(WORKED_HEADERS());
    styleHeader(wSheet.getRange(1, 1, 1, WORKED_HEADERS().length));
    wSheet.setFrozenRows(1);
    setColumnWidths(wSheet, [60, 120, 180, 220, 140, 180, 140, 120, 200]);
  }
  
  // "Статистика"
  if (!ss.getSheetByName(STATS_SHEET)) {
    const stSheet = ss.insertSheet(STATS_SHEET);
    stSheet.appendRow(STATS_HEADERS());
    styleHeader(stSheet.getRange(1, 1, 1, STATS_HEADERS().length));
    stSheet.setFrozenRows(1);
    setColumnWidths(stSheet, [120, 90, 80, 80, 80, 90, 90, 90, 80]);
  }
  
  return "✅ Листы созданы: " + [ADMIN_SHEET, TIER_SHEET, WORKED_SHEET, STATS_SHEET].join(", ");
}

// выгрузка JSON в файл
function saveJSONToFile() {
  const json = exportJSON();
  if (json.includes('error')) return json;
  
  const file = DriveApp.createFile('admin_crm_data.json', json, 'application/json');
  return "✅ JSON сохранён: " + file.getUrl();
}

// запуск по меню
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("_ВЛАД Admin v2.0")
    .addItem("Создать листы", "initSystem")
    .addItem("Синхронизировать", "syncAllSheets")
    .addItem("Сортировать", "sortTierList")
    .addItem("Экспорт JSON", "saveJSONToFile")
    .addToUi();
}
