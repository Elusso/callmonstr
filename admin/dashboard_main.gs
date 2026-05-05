/**
 * ADMIN CRM — Офис Молодость (v2.4)
 * =================================
 * Минималистичная CRM с базовым оформлением.
 * Считывает лиды со всех сотрудников.
 */

const ADMIN_SPREADSHEET_ID = "1AUCWikqIhAGxXKVvTm1SLJjmLoUh0Y_c7Sry8X8cX8cXos";

const EMPLOYEES = [
  {name: "Тёмыч",  id: "1VCVAZhTl4cv9T1J4AyzknYqekMy6ZEymlQRkvOB4yJE"},
  {name: "Влад",   id: "1Lt9BmIVShNFserfYacxII6WjGPDn5ObNeiOo9z63oI8"},
  {name: "Соня",   id: "1U2uq6xhVXxcUvN3bTwj0eleQZ5mUm4c4OCpwN-deqoA"},
  {name: "Костян", id: "1CcaWPBvdPZ5WwegxDKOQhxljvMHwPZnYLgpl05QVB1c"},
  {name: "Денишк", id: "1pYRyigxMNSmrqr92RZ9-I5rbehErgX-PSiRbhn4Z4qU"}
];

const LEADS_SHEET = "Все лиды";
const STATS_SHEET = "Статистика";
const ACTIVES_SHEET = "Активные лиды";

const ACTIVE_STATUSES = ["📝 Заявка", "🎫 Ожидает билеты", "🚗 Ожидает выезда", "🚀 В пути", "🏛️ В военкомате", "🔍 На проверке", "✅ ПОДПИСАН"];

// Получение всех лидов
function getAllLeads() {
  let all = [];
  EMPLOYEES.forEach(emp => {
    try {
      const ss = SpreadsheetApp.openById(emp.id);
      const sheet = ss.getSheets()[0];
      if (!sheet) return;
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return;
      
      for (let i = 1; i < data.length; i++) {
        const r = data[i];
        if (r.length < 8) continue;
        all.push({
          name: emp.name,
          date: r[0],
          vacancy: r[1],
          city: r[2],
          fio: r[3],
          phone: r[4],
          age: r[5],
          status: r[6],
          notes: r[7]
        });
      }
    } catch (e) {
      Logger.log(`⚠️ ${emp.name}: ${e}`);
    }
  });
  return all;
}

// Создание листов
function initSystem() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  
  // "Все лиды" - базовый лист
  let sheet = ss.getSheetByName(LEADS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LEADS_SHEET);
    sheet.appendRow(["ID", "Сотрудник", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"]);
    headerStyle(sheet);
  }
  
  // "Статистика"
  sheet = ss.getSheetByName(STATS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(STATS_SHEET);
    sheet.appendRow(["Дата", "Всего", "Подписано", "Отказ", "Свяьь", "Активные", "Успех %"]);
    headerStyle(sheet);
  }
  
  // "Активные лиды"
  sheet = ss.getSheetByName(ACTIVES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ACTIVES_SHEET);
    sheet.appendRow(["ID", "Сотрудник", "Дата", "ФИО", "Телефон", "Статус", "Заметки"]);
    headerStyle(sheet);
  }
  
  return "✅ Листы созданы";
}

function headerStyle(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  range.setBackground("#1a1a1a").setFontColor("#00ff88").setFontWeight("bold");
}

// Синхронизация
function syncAllSheets() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const allLeads = getAllLeads();
  
  // 1. Все лиды
  const sheet = ss.getSheetByName(LEADS_SHEET);
  if (!sheet) return "❌ Не найден лист 'Все лиды'";
  
  sheet.clear();
  sheet.appendRow(["ID", "Сотрудник", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"]);
  headerStyle(sheet);
  
  if (allLeads.length > 0) {
    allLeads.sort((a, b) => {
      const dA = a.date ? new Date(a.date).getTime() : 0;
      const dB = b.date ? new Date(b.date).getTime() : 0;
      return dB - dA;
    });
    
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
    
    // Write without validation - just data
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    
    // Simple alternating row colors
    for (let i = 0; i < data.length; i++) {
      const row = sheet.getRange(2 + i, 1, 1, data[0].length);
      row.setBackground(i % 2 === 0 ? "#2d2d2d" : "#1a1a1a");
    }
  }
  
  // 2. Статистика
  const statsSheet = ss.getSheetByName(STATS_SHEET);
  if (statsSheet) {
    statsSheet.clear();
    statsSheet.appendRow(["Дата", "Всего", "Подписано", "Отказ", "Свяьь", "Активные", "Успех %"]);
    headerStyle(statsSheet);
    
    let total = allLeads.length;
    let connected = 0, refused = 0, contact = 0, activeCount = 0;
    
    allLeads.forEach(l => {
      const s = l.status || "";
      if (s.includes("✅")) connected++;
      if (s.includes("🔴") || s.includes("⚫") || s.includes("❌")) refused++;
      if (s.includes("💬") || s.includes("🟡")) contact++;
      if (ACTIVE_STATUSES.indexOf(s) >= 0) activeCount++;
    });
    
    const successRate = total > 0 ? Math.round(100 * connected / total) : 0;
    statsSheet.appendRow([new Date(), total, connected, refused, contact, activeCount, successRate + "%"]);
  }
  
  // 3. Активные лиды
  const activesSheet = ss.getSheetByName(ACTIVES_SHEET);
  if (activesSheet) {
    activesSheet.clear();
    activesSheet.appendRow(["ID", "Сотрудник", "Дата", "ФИО", "Телефон", "Статус", "Заметки"]);
    headerStyle(activesSheet);
    
    const actives = allLeads.filter(l => ACTIVE_STATUSES.indexOf(l.status || "") >= 0);
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
      
      activesSheet.getRange(2, 1, data.length, data[0].length).setValues(data);
      
      for (let i = 0; i < data.length; i++) {
        const row = activesSheet.getRange(2 + i, 1, 1, data[0].length);
        row.setBackground(i % 2 === 0 ? "#2d2d2d" : "#1a1a1a");
      }
    }
  }
  
  return "✅ Синхронизация: " + allLeads.length + " лидов";
}

// Меню
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("_Офис Молодость Admin v2.4")
    .addItem("Создать листы", "initSystem")
    .addItem("Синхронизировать", "syncAllSheets")
    .addToUi();
}
