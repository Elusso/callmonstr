#!/usr/bin/env node
// callmonstr_v9_gas.js - Generated Google Apps Script
// Dark Theme: #1a1a1a, Neon Green #00ff88, Comfortaa font

const CONFIG = {
  "theme": {
    "bgDeep": "#1a1a1a",
    "bgMid": "#2d2d2d",
    "accent": "#00ff88",
    "accentRed": "#ff3f34",
    "warning": "#ffa502",
    "textMain": "#e0e0e0",
    "font": "Comfortaa",
    "fontSize": 14,
    "rowHeight": 30,
    "headerHeight": 40
  },
  "tabs": {
    "instruction": "📒 Инструкция",
    "allLeads": "Всё лиды",
    "active": "🎯 Активные",
    "employees": "👥 Сотрудники",
    "stats": "📊 Статистика",
    "history": "📝 История",
    "archive": "🗄 Архив",
    "log": "📝 Лог"
  },
  "columns": {
    "allLeads": [
      "Дата",
      "Вакансия",
      "Город",
      "ФИО",
      "Телефон",
      "Возраст",
      "Статус",
      "Заметки"
    ],
    "active": [
      "Дата",
      "Сотрудник",
      "ФИО",
      "Телефон",
      "Вакансия",
      "Город",
      "Статус",
      "Выезд",
      "Приезд",
      "Заметки"
    ],
    "employees": [
      "Сотрудник",
      "ID",
      "Ссылка",
      "Статус",
      "Дата",
      "Лидов",
      "Дозвон %"
    ],
    "stats": [
      "Период",
      "Сотрудник",
      "Всего",
      "Дозвон",
      "НД",
      "Подписано",
      "Комиссовано",
      "%",
      "Возраст",
      "График"
    ]
  },
  "statuses": {
    "valid": [
      "⚪️ Новый",
      "🔴 НД",
      "🤔 ДУМ",
      "🟡 ПЕРЕЗВОНИТЬ",
      "💬 СВЯЗЬ МЕССЕНДЖЕР",
      "✅ ПОДПИСАН",
      "✅ Подписан",
      "⚫️ ОТКАЗ",
      "❌ Отказ",
      "📝 Заявка",
      "🎫 Ожидает билеты",
      "🚗 Ожидает выезда",
      "🚀 В пути",
      "🏛 В военкомате",
      "🔍 На проверке",
      "🎗️ Комиссован",
      "🎗 Комиссован"
    ]
  },
  "sparklines": {
    "chars": "▁▂▃▄▅▆▇█"
  }
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 CallMonstr v9')
    .addItem('🔄 Синхронизация', 'syncAllData')
    .addItem('📊 Обновить статистику', 'updateStats')
    .addItem('🚀 Шаттл НД', 'shuffleOldND')
    .addItem('💾 Бэкап', 'backupSpreadsheet')
    .addToUi();
}

function initFullSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheets().map(s => s.getName());
  
  const sheetConfigs = [
    {name: CONFIG.tabs.instruction, cols: 4},
    {name: CONFIG.tabs.allLeads, cols: 8},
    {name: CONFIG.tabs.active, cols: 10},
    {name: CONFIG.tabs.employees, cols: 7},
    {name: CONFIG.tabs.stats, cols: 10},
    {name: CONFIG.tabs.history, cols: 7},
    {name: CONFIG.tabs.archive, cols: 7},
    {name: CONFIG.tabs.log, cols: 4}
  ];
  
  sheetConfigs.forEach(config => {
    if (!existing.includes(config.name)) {
      ss.insertSheet(config.name);
    }
    const ws = ss.getSheetByName(config.name);
    ws.clear();
    ws.getRange(1, 1, 1, config.cols).setFontFamily(CONFIG.theme.font)
     .setFontSize(CONFIG.theme.fontSize + 2)
     .setFontWeight("bold")
     .setBackground(CONFIG.theme.bgDeep)
     .setFontColor(CONFIG.theme.accent);
    ws.getRowDimension(1).setHeight(CONFIG.theme.headerHeight);
  });
}

function setupInstructionDark(ws) {
  ws.getRange("B1").setValue("CALLMONSTR V9.0");
  ws.getRange("B2").setValue("ТЕМНЫЙ ТЕРМИНАЛ v9.0");
  ws.getRange("B1:C2").merge().setFontFamily(CONFIG.theme.font).setFontSize(18);
  ws.getRowDimension(1).setHeight(50);
  ws.getRowDimension(2).setHeight(40);
}

function syncAllData() {
  // Auto-generated sync function
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allLeads = ss.getSheetByName(CONFIG.tabs.allLeads);
  const activeSheet = ss.getSheetByName(CONFIG.tabs.active);
  const historySheet = ss.getSheetByName(CONFIG.tabs.history);
  
  allLeads.getRange("A2:ZZ").clearContent();
  activeSheet.getRange("A2:ZZ").clearContent();
  historySheet.getRange("A2:ZZ").clearContent();
}

function updateStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statsSheet = ss.getSheetByName(CONFIG.tabs.stats);
  statsSheet.getRange("A2:ZZ").clearContent();
}

function shuffleOldND() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.tabs.allLeads);
}

function backupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = ss.getName() + " - Backup " + new Date().toISOString().split("T")[0];
  DriveApp.getRootFolder().createFile(name, ss.getBlob().getDataAsString(), "application/vnd.google-apps.spreadsheet");
}

function formatSparkline(values) {
  const chars = CONFIG.sparklines.chars;
  if (!values || values.length === 0) return chars.repeat(12);
  const min = Math.min(...values), max = Math.max(...values);
  if (min === max) return "▄".repeat(12);
  return values.map(v => {
    const idx = Math.floor(((v - min) / (max - min)) * (chars.length - 1));
    return chars[Math.max(0, Math.min(chars.length - 1, idx))];
  }).join("");
}

function isValidStatus(status) {
  return CONFIG.statuses.valid.includes(status.trim());
}

function isND(status) {
  return status.toLowerCase().includes("нд");
}

function isPodpisan(status) {
  return status.toLowerCase().includes("подписан");
}

function isKomissovan(status) {
  return status.toLowerCase().includes("комиссован");
}
