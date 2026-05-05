/**
 * ADMIN CRM — Office Молодость v3.0
 * ==================================
 * Сбор лидов со всех сотрудников, статистика, активные лиды.
 * Парсит A-H колонки (0-7): Дата, Вакансия, Город, ФИО, Телефон, Возраст, Статус, Заметки
 */

const ADMIN_SPREADSHEET_ID = "1AUCWikqIhAGxXKVvTm1SLJjmLoUh0Y_c7Sry8X8cXos";

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

const ACTIVE_STATUSES = ["📝 Заявка", "🎫 Ожидает билеты", "🚗 Ожидает выезда", "🚀 В пути", "🏛️ В военкомате", "🔍 На проверке", "✅ ПОДПИСАН", "✅ Подписан"];

// Получение лидов из одной таблицы
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
        date: r[0] || "",
        vacancy: r[1] || "",
        city: r[2] || "",
        fio: r[3] || "",
        phone: r[4] || "",
        age: r[5] || "",
        status: r[6] || "",
        notes: r[7] || ""
      });
    }
    return rows;
  } catch (e) {
    Logger.log("Ошибка " + emp.name + ": " + e);
    return [];
  }
}

// Сбор всех лидов
function getAllLeads() {
  let all = [];
  EMPLOYEES.forEach(emp => {
    const leads = readEmployeeSheet(emp);
    all = all.concat(leads);
  });
  Logger.log("Считано лидов: " + all.length);
  return all;
}

// Создание листов
function initSystem() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  
  const leadsSheet = ss.getSheetByName(LEADS_SHEET) || ss.insertSheet(LEADS_SHEET);
  leadsSheet.clear();
  leadsSheet.appendRow(["ID", "Сотрудник", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"]);
  headerStyle(leadsSheet);
  leadsSheet.setColumnWidths([50, 120, 100, 180, 120, 220, 140, 70, 120, 180]);
  
  const statsSheet = ss.getSheetByName(STATS_SHEET) || ss.insertSheet(STATS_SHEET);
  statsSheet.clear();
  statsSheet.appendRow(["Дата", "Всего", "Подписано", "Отказ", "Свяьь", "Активные", "Успех %"]);
  headerStyle(statsSheet);
  
  const activesSheet = ss.getSheetByName(ACTIVES_SHEET) || ss.insertSheet(ACTIVES_SHEET);
  activesSheet.clear();
  activesSheet.appendRow(["ID", "Сотрудник", "Дата", "ФИО", "Телефон", "Статус", "Заметки"]);
  headerStyle(activesSheet);
  
  return "✅ Листы созданы: Все лиды, Статистика, Активные лиды";
}

function headerStyle(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  range.setBackground("#1a1a1a").setFontColor("#00ff88").setFontWeight("bold").setFontSize(12);
}

// Синхронизация
function syncAllSheets() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const allLeads = getAllLeads();
  
  // 1. Все лиды
  const leadsSheet = ss.getSheetByName(LEADS_SHEET);
  if (!leadsSheet) return "❌ Не найден лист 'Все лиды'";
  
  leadsSheet.clearContents();
  leadsSheet.appendRow(["ID", "Сотрудник", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"]);
  headerStyle(leadsSheet);
  
  // Сортировка по дате (новее сверху)
  allLeads.sort((a, b) => {
    const dA = a.date ? new Date(a.date).getTime() : 0;
    const dB = b.date ? new Date(b.date).getTime() : 0;
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
    
    // Чередующаяся заливка
    for (let i = 0; i < data.length; i++) {
      leadsSheet.getRange(2 + i, 1, 1, data[0].length).setBackground(i % 2 === 0 ? "#2d2d2d" : "#1a1a1a");
    }
  }
  
  // 2. Статистика
  const statsSheet = ss.getSheetByName(STATS_SHEET);
  if (statsSheet) {
    statsSheet.clearContents();
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
    activesSheet.clearContents();
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
      
      const range = activesSheet.getRange(2, 1, data.length, data[0].length);
      range.setValues(data);
      
      for (let i = 0; i < data.length; i++) {
        activesSheet.getRange(2 + i, 1, 1, data[0].length).setBackground(i % 2 === 0 ? "#2d2d2d" : "#1a1a1a");
      }
    }
  }
  
  return "✅ Синхронизация: " + allLeads.length + " лидов";
}

// Меню
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("_Офис Молодость Admin v3.0")
    .addItem("Создать листы", "initSystem")
    .addItem("Синхронизировать", "syncAllSheets")
    .addToUi();
}
