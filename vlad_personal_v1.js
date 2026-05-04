/**
 * CallMonstr v1.0 — Vlad Personal Sheet
 * Синхронизация Tier-list с основным листом лидов
 *
 * Листы:
 * - "Все лиды" — основная таблица (вакансия, ФИО, телефон, статус, тир, комментарий)
 * - "Отработать" — НД лиды после шаттла (от других сотрудников)
 * - "Tier-list" — синхронизированный с "Все лиды", отсортирован по тиру, с датой напоминания
 *
 * Функции:
 * - initSystem() — создание всех листов
 * - syncTierList() — синхронизация "Tier-list" с "Все лиды"
 * - sortTierList() — сортировка по тиру (сверху ↓, 1→5)
 * - getReminderDate() — дата напоминания (через 3 дня от даты добавления)
 */

const EMPLOYEE_ID = "1Lt9BmIVShNFserfYacxII6WjGPDn5ObNeiOo9z63oI8"; // Влад
const ADMIN_SPREADSHEET_ID = "1AUCWikqIhAGxXKVvTm1SLJjmLoUh0Y_c7Sry8X8cXos";

const MAIN_HEADERS = ["Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Тир", "Комментарий", "Дата напоминания"];
const WORKED_COL_HEADERS = ["Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Статус", "Комментарий"];
const TIER_LIST_HEADERS = ["Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Тир", "Комментарий", "Дата напоминания"];

const VALID_TIERS = [1, 2, 3, 4, 5];

function initSystem() {
  const ss = SpreadsheetApp.openById(EMPLOYEE_ID);
  
  // Check if "Все лиды" exists, create if not
  let mainSheet = ss.getSheetByName("Все лиды");
  if (!mainSheet) {
    mainSheet = ss.insertSheet("Все лиды");
    mainSheet.appendRow(MAIN_HEADERS);
    const r = mainSheet.getRange(1, 1, 1, MAIN_HEADERS.length);
    r.setBackground("#1a1a1a").setFontColor("#00ff88").setFontFamily("Comfortaa").setFontWeight("bold").setFontSize(11);
    mainSheet.setFrozenRows(1);
    [120, 180, 140, 220, 140, 80, 120, 60, 200, 150].forEach((w, i) => { try { mainSheet.setColumnWidth(i + 1, w); } catch {} });
    mainSheet.setRowHeight(1, 40);
  }
  
  // Create "Отработать" sheet
  let workedSheet = ss.getSheetByName("Отработать");
  if (!workedSheet) {
    workedSheet = ss.insertSheet("Отработать");
    workedSheet.appendRow(WORKED_COL_HEADERS);
    const r = workedSheet.getRange(1, 1, 1, WORKED_COL_HEADERS.length);
    r.setBackground("#1a1a1a").setFontColor("#00ff88").setFontFamily("Comfortaa").setFontWeight("bold").setFontSize(11);
    workedSheet.setFrozenRows(1);
    [120, 180, 220, 140, 180, 140, 120, 200].forEach((w, i) => { try { workedSheet.setColumnWidth(i + 1, w); } catch {} });
    workedSheet.setRowHeight(1, 40);
  }
  
  // Create "Tier-list" sheet
  let tierSheet = ss.getSheetByName("Tier-list");
  if (!tierSheet) {
    tierSheet = ss.insertSheet("Tier-list");
    tierSheet.appendRow(TIER_LIST_HEADERS);
    const r = tierSheet.getRange(1, 1, 1, TIER_LIST_HEADERS.length);
    r.setBackground("#1a1a1a").setFontColor("#00ff88").setFontFamily("Comfortaa").setFontWeight("bold").setFontSize(11);
    tierSheet.setFrozenRows(1);
    [120, 180, 140, 220, 140, 80, 120, 60, 200, 150].forEach((w, i) => { try { tierSheet.setColumnWidth(i + 1, w); } catch {} });
    tierSheet.setRowHeight(1, 40);
  }
  
  return "✅ Система создана: Все лиды, Отработать, Tier-list";
}

function syncTierList() {
  const ss = SpreadsheetApp.openById(EMPLOYEE_ID);
  const mainSheet = ss.getSheetByName("Все лиды");
  const tierSheet = ss.getSheetByName("Tier-list");
  
  if (!mainSheet || !tierSheet) {
    return "❌ Не найдены листы. Выполните initSystem()";
  }
  
  // Get data from "Все лиды" (skip header)
  const mainData = mainSheet.getDataRange().getValues();
  if (mainData.length <= 1) {
    return "⏸ Нет данных в 'Все лиды'";
  }
  
  // Clear "Tier-list" except header
  if (tierSheet.getLastRow() > 1) {
    tierSheet.deleteRows(2, tierSheet.getLastRow() - 1);
  }
  
  // Extract relevant columns from "Все лиды" and sync to "Tier-list"
  const newData = [];
  for (let i = 1; i < mainData.length; i++) {
    const row = mainData[i];
    if (row.length < 9) continue; // Need: date, vacancy, city, fio, phone, age, status, tier, comment
    
    const tier = row[7] !== undefined ? row[7] : "";
    const comment = row[8] !== undefined ? row[8] : "";
    
    // Build tier-list row from main
    let tgt = [row[0], row[1], row[2], row[3], row[4], row[5], row[6], tier, comment];
    
    // Calculate reminder date (3 days from lead date)
    const leadDate = row[0];
    if (leadDate instanceof Date) {
      const remDate = new Date(leadDate.getTime() + 3 * 24 * 60 * 60 * 1000);
      tgt.push(remDate);
    } else {
      tgt.push("");
    }
    
    newData.push(tgt);
  }
  
  if (newData.length > 0) {
    tierSheet.getRange(tierSheet.getLastRow() + 1, 1, newData.length, newData[0].length).setValues(newData);
    
    // Apply color coding for tiers (top to bottom)
    for (let i = 0; i < newData.length; i++) {
      const tier = parseInt(newData[i][7]);
      const cell = tierSheet.getRange(i + 2, 8); // tier column
      if (tier === 1) cell.setBackground("#ff3f34"); // red
      else if (tier === 2) cell.setBackground("#ffa502"); // orange
      else if (tier === 3) cell.setBackground("#ffff00"); // yellow
      else if (tier === 4) cell.setBackground("#add8e6"); // light blue
      else if (tier === 5) cell.setBackground("#39ff14"); // green
      else cell.setBackground("#ffffff");
    }
  }
  
  // Sort by tier (top to bottom: 1,2,3,4,5)
  sortTierList();
  
  return "✅ Синхронизация завершена. " + newData.length + " записей";
}

function sortTierList() {
  const ss = SpreadsheetApp.openById(EMPLOYEE_ID);
  const tierSheet = ss.getSheetByName("Tier-list");
  
  if (!tierSheet) return "❌ Tier-list не найден";
  if (tierSheet.getLastRow() <= 1) return "⏸ Нет данных для сортировки";
  
  const data = tierSheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);
  
  // Sort by tier (numeric), then by date (descending)
  rows.sort((a, b) => {
    const tierA = parseInt(a[7]) || 999;
    const tierB = parseInt(b[7]) || 999;
    if (tierA !== tierB) return tierA - tierB; // 1 before 2, etc.
    // Same tier: newer date first
    const dateA = new Date(a[0]).getTime() || 0;
    const dateB = new Date(b[0]).getTime() || 0;
    return dateB - dateA;
  });
  
  // Clear and rewrite sorted data
  if (tierSheet.getLastRow() > 1) {
    tierSheet.deleteRows(2, tierSheet.getLastRow() - 1);
  }
  
  if (rows.length > 0) {
    tierSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    
    // Re-apply tier colors after sort
    for (let i = 0; i < rows.length; i++) {
      const tier = parseInt(rows[i][7]);
      const cell = tierSheet.getRange(i + 2, 8);
      if (tier === 1) cell.setBackground("#ff3f34");
      else if (tier === 2) cell.setBackground("#ffa502");
      else if (tier === 3) cell.setBackground("#ffff00");
      else if (tier === 4) cell.setBackground("#add8e6");
      else if (tier === 5) cell.setBackground("#39ff14");
    }
  }
  
  return "⬆️ Сортировка по тиру завершена";
}

function setupTriggers() {
  // Delete existing triggers
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  
  // Daily sync at 9:00
  ScriptApp.newTrigger("syncTierList").timeBased().atHour(9).everyDays(1).create();
  
  return "✅ Триггеры настроены";
}

function showMenu() {
  return [
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "   Влад — CallMonstr v1.0",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "init()    — Создать листы",
    "sync()    — Синхронизировать Tier-list",
    "sort()    — Сортировать по тиру",
    "triggers() — Настроить триггеры",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
  ].join("\n");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("_ВЛАД v1.0")
    .addItem("Создать листы", "initSystem")
    .addItem("Синхронизировать", "syncTierList")
    .addItem("Сортировать", "sortTierList")
    .addItem("Триггеры", "setupTriggers")
    .addItem("Меню", "showMenu")
    .addToUi();
}

function main(args = []) {
  if (!args || args.length === 0) return showMenu();
  const cmd = (args[0] || "").toLowerCase();
  switch (cmd) {
    case "init": return initSystem();
    case "sync": return syncTierList();
    case "sort": return sortTierList();
    case "triggers": return setupTriggers();
    default: return showMenu();
  }
}
