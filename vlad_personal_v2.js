/**
 * CallMonstr v2.0 — Vlad Personal Sheet (SS/S/A)
 * Двусторонняя синхронизация Tier-list с "Все лиды"
 *
 * Tier logic:
 *   SS = "подписан + готов + без рисков" (top candidates)
 *   S  = "готов, но с нюансами" (KPI < 80%, доп. проверка)
 *   A  = "в процессе конверсии" (связь есть, не подписан)
 *
 * Функции:
 * - initSystem() — создание всех листов
 * - syncTierList() — синхронизация "Tier-list" с "Все лиды"
 * - sortTierList() — сортировка по тиру (SS → S → A)
 * - onEdit(e) — двусторонняя синхронизация при редактировании
 * - setupTriggers() — создание time-based триггеров
 */

const EMPLOYEE_ID = "1Lt9BmIVShNFserfYacxII6WjGPDn5ObNeiOo9z63oI8";
const MAIN_SHEET = "Все лиды";
const TIER_SHEET = "Tier-list";
const WORKED_SHEET = "Отработать";

// status→tier mapping
const STATUS_TO_TIER = {
  // SS: подписан + готов + без рисков
  "✅ Подписан": "SS",
  "✅ Подписаны": "SS",
  "✅ Подписан": "SS",
  "✅ ПОДПИСАН": "SS",
  "🎗️ Комиссован": "SS",
  // S: готов, но с нюансами
  "🔴 НД": "S",
  "🤔 ДУМ": "S",
  "⚫ ОТКАЗ": "S",
  "❌ Отказ": "S",
  // A: в процессе конверсии
  "🟡 ПЕРЕЗВОНИТЬ": "A",
  "💬 СВЯЗЬ МЕССЕНДЖЕР": "A",
  "🎫 Ожидает билеты": "A",
  "🚗 Ожидает выезда": "A",
  "🚀 В пути": "A",
  "🏛️ В военкомате": "A",
  "🔍 На проверке": "A",
  "📝 Заявка": "A",
  // unknown → A
  "default": "A"
};

const MAIN_COLS = { date: 0, vacancy: 1, city: 2, fio: 3, phone: 4, age: 5, status: 6, tier: 7, comment: 8, reminder: 9 };
const TIER_COLS = { date: 0, vacancy: 1, city: 2, fio: 3, phone: 4, age: 5, status: 6, tier: 7, comment: 8, reminder: 9 };

function initSystem() {
  const ss = SpreadsheetApp.openById(EMPLOYEE_ID);
  
  // "Все лиды" — main sheet
  let mainSheet = ss.getSheetByName(MAIN_SHEET);
  if (!mainSheet) {
    mainSheet = ss.insertSheet(MAIN_SHEET);
    mainSheet.appendRow(MAIN_HEADERS());
    styleHeader(mainSheet.getRange(1, 1, 1, MAIN_HEADERS().length));
    mainSheet.setFrozenRows(1);
    setColumnWidths(mainSheet, [120, 180, 140, 220, 140, 80, 120, 60, 200, 150]);
  }
  
  // "Отработать"
  let workedSheet = ss.getSheetByName(WORKED_SHEET);
  if (!workedSheet) {
    workedSheet = ss.insertSheet(WORKED_SHEET);
    workedSheet.appendRow(WORKED_HEADERS());
    styleHeader(workedSheet.getRange(1, 1, 1, WORKED_HEADERS().length));
    workedSheet.setFrozenRows(1);
    setColumnWidths(workedSheet, [120, 180, 220, 140, 180, 140, 120, 200]);
  }
  
  // "Tier-list"
  let tierSheet = ss.getSheetByName(TIER_SHEET);
  if (!tierSheet) {
    tierSheet = ss.insertSheet(TIER_SHEET);
    tierSheet.appendRow(TIER_HEADERS());
    styleHeader(tierSheet.getRange(1, 1, 1, TIER_HEADERS().length));
    tierSheet.setFrozenRows(1);
    setColumnWidths(tierSheet, [120, 180, 140, 220, 140, 80, 120, 60, 200, 150]);
  }
  
  return "✅ Листы созданы: " + [MAIN_SHEET, WORKED_SHEET, TIER_SHEET].join(", ");
}

function syncTierList() {
  const ss = SpreadsheetApp.openById(EMPLOYEE_ID);
  const mainSheet = ss.getSheetByName(MAIN_SHEET);
  const tierSheet = ss.getSheetByName(TIER_SHEET);
  
  if (!mainSheet || !tierSheet) return "❌ Не найдены листы";
  
  const mainData = mainSheet.getDataRange().getValues();
  if (mainData.length <= 1) return "⏸ Нет данных в 'Все лиды'";
  
  // Clear tier sheet (preserve header)
  if (tierSheet.getLastRow() > 1) tierSheet.deleteRows(2, tierSheet.getLastRow() - 1);
  
  const newData = [];
  for (let i = 1; i < mainData.length; i++) {
    const row = mainData[i];
    if (row.length < 10) continue;
    
    // Find phone in tier sheet, check if exists
    const phone = row[MAIN_COLS.phone];
    if (phone) {
      const tgt = [
        row[MAIN_COLS.date], row[MAIN_COLS.vacancy], row[MAIN_COLS.city],
        row[MAIN_COLS.fio], row[MAIN_COLS.phone], row[MAIN_COLS.age],
        row[MAIN_COLS.status], getStatusTier(row[MAIN_COLS.status]), row[MAIN_COLS.comment]
      ];
      
      // Reminder date: today (not from lead date)
      const remDate = new Date();
      tgt.push(remDate);
      newData.push(tgt);
    }
  }
  
  if (newData.length > 0) {
    tierSheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
    applyTierColors(tierSheet, 2, newData.length);
  }
  
  sortTierList();
  return "✅ Синхронизация: " + newData.length + " записей";
}

function sortTierList() {
  const ss = SpreadsheetApp.openById(EMPLOYEE_ID);
  const tierSheet = ss.getSheetByName(TIER_SHEET);
  if (!tierSheet) return "❌ Tier-list не найден";
  
  const data = tierSheet.getDataRange().getValues();
  if (data.length <= 1) return "⏸ Нет данных";
  
  const header = data[0];
  const rows = data.slice(1);
  
  // SS → S → A ordering
  rows.sort((a, b) => {
    const tierA = getTierRank(a[TIER_COLS.tier]);
    const tierB = getTierRank(b[TIER_COLS.tier]);
    if (tierA !== tierB) return tierA - tierB;
    // Same tier: newer date first
    const dateA = parseDate(a[TIER_COLS.date]);
    const dateB = parseDate(b[TIER_COLS.date]);
    return dateB - dateA;
  });
  
  if (tierSheet.getLastRow() > 1) tierSheet.deleteRows(2, tierSheet.getLastRow() - 1);
  if (rows.length > 0) {
    tierSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    applyTierColors(tierSheet, 2, rows.length);
  }
  
  return "⬆️ Сортировка: " + rows.length + " записей";
}

function onEdit(e) {
  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    
    // Only "Все лиды" and "Tier-list" trigger sync
    if (sheetName === MAIN_SHEET || sheetName === TIER_SHEET) {
      const col = e.range.getColumn();
      // Only TIER (8) and COMMENT (9) columns trigger sync
      if (col === MAIN_COLS.tier + 1 || col === MAIN_COLS.comment + 1 || 
          col === TIER_COLS.tier + 1 || col === TIER_COLS.comment + 1) {
        
        syncTierList();
        return;
      }
    }
  } catch (err) {
    Logger.log("onEdit error: " + err);
  }
}

function setupTriggers() {
  // Remove existing triggers
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  
  // Daily sync at 9:00
  ScriptApp.newTrigger("syncTierList")
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
  
  // On-edit trigger for bidirectional sync
  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(SpreadsheetApp.openById(EMPLOYEE_ID))
    .onEdit()
    .create();
  
  return "✅ Триггеры настроены";
}

// Helpers
function MAIN_HEADERS() { return ["Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Тир", "Комментарий", "Дата напоминания"]; }
function TIER_HEADERS() { return ["Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Тир", "Комментарий", "Дата напоминания"]; }
function WORKED_HEADERS() { return ["Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Статус", "Комментарий"]; }

function styleHeader(range) {
  range.setBackground("#1a1a1a").setFontColor("#00ff88").setFontFamily("Comfortaa").setFontWeight("bold").setFontSize(11);
}

function setColumnWidths(sheet, widths) {
  widths.forEach((w, i) => { try { sheet.setColumnWidth(i + 1, w); } catch {} });
  if (sheet.getLastRow() >= 1) sheet.setRowHeight(1, 40);
}

function applyTierColors(sheet, startRow, rowCount) {
  for (let r = 0; r < rowCount; r++) {
    const tier = sheet.getRange(startRow + r, TIER_COLS.tier + 1).getValue();
    const cell = sheet.getRange(startRow + r, TIER_COLS.tier + 1);
    
    if (tier === "SS") cell.setBackground("#ff3f34");      // красный (top tier)
    else if (tier === "S") cell.setBackground("#ffa502");   // оранжевый (second)
    else if (tier === "A") cell.setBackground("#ffff00");   // жёлтый (active)
    else cell.setBackground("#ffffff");
  }
}

function getStatusTier(status) {
  if (!status) return STATUS_TO_TIER["default"];
  return STATUS_TO_TIER[status.trim()] || STATUS_TO_TIER["default"];
}

function getTierRank(tier) {
  if (tier === "SS") return 1;
  if (tier === "S") return 2;
  if (tier === "A") return 3;
  return 999;
}

function parseDate(val) {
  if (!val) return 0;
  const d = new Date(val);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function showMenu() {
  return [
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "   Влад — CallMonstr v2.0",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "init()  — Создать листы",
    "sync()  — Синхронизировать Tier-list",
    "sort()  — Сортировать по тиру",
    "triggers() — Настроить триггеры",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
  ].join("\n");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("_ВЛАД v2.0")
    .addItem("Создать листы", "initSystem")
    .addItem("Синхронизировать", "syncTierList")
    .addItem("Сортировать", "sortTierList")
    .addItem("Триггеры", "setupTriggers")
    .addToUi();
}