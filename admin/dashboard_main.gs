/**
 * ADMIN CRM — Office Молодость v3.1
 * ==================================
 * ПОЛНАЯ РАБОТОСПОСОБНАЯ ВЕРСИЯ
 * Сбор лидов со всех сотрудников, статистика, активные лиды.
 * Парсит A-H колонки (0-7): Дата, Вакансия, Город, ФИО, Телефон, Возраст, Статус, Заметки
 */

const ADMIN_SPREADSHEET_ID = "1AUCWikqIhAGxXKVvTm1SLJjmLoUh0Y_c7Sry8X8cXos";

// 5 СОТРУДНИКОВ (из памяти)
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

// БЕЗОПАСНОЕ获得ение данных
function safeReadSheet(emp) {
  try {
    const ss = SpreadsheetApp.openById(emp.id);
    const sheet = ss.getSheets()[0];
    if (!sheet) {
      Logger.log("⚠️ Лист не найден у " + emp.name);
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log("⚠️ Нет данных у " + emp.name);
      return [];
    }
    
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      // Проверка на минимум 8 колонок
      if (r.length < 8) {
        Logger.log("⚠️ Слишком мало колонок у " + emp.name + " строка " + (i + 1));
        continue;
      }
      
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
    Logger.log("✅ " + emp.name + ": " + rows.length + " лидов");
    return rows;
  } catch (e) {
    Logger.log("❌ Ошибка " + emp.name + ": " + e);
    return [];
  }
}

// Сбор всех лидов
function getAllLeads() {
  let all = [];
  EMPLOYEES.forEach(emp => {
    const leads = safeReadSheet(emp);
    all = all.concat(leads);
  });
  Logger.log("📊 ВСЕГО ЛИДОВ: " + all.length);
  return all;
}

// Создание листов
function initSystem() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  
  // "Все лиды"
  let sheet = ss.getSheetByName(LEADS_SHEET);
  if (!sheet) sheet = ss.insertSheet(LEADS_SHEET);
  sheet.clear();
  sheet.appendRow(["ID", "Сотрудник", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"]);
  headerStyle(sheet, 10);
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 220);
  sheet.setColumnWidth(7, 140);
  sheet.setColumnWidth(8, 70);
  sheet.setColumnWidth(9, 120);
  sheet.setColumnWidth(10, 180);
  sheet.setFrozenRows(1);
  
  // "Статистика"
  sheet = ss.getSheetByName(STATS_SHEET);
  if (!sheet) sheet = ss.insertSheet(STATS_SHEET);
  sheet.clear();
  sheet.appendRow(["Дата", "Всего", "Подписано", "Отказ", "Свяьь", "Активные", "Успех %"]);
  headerStyle(sheet, 7);
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 80);
  sheet.setFrozenRows(1);
  
  // "Активные лиды"
  sheet = ss.getSheetByName(ACTIVES_SHEET);
  if (!sheet) sheet = ss.insertSheet(ACTIVES_SHEET);
  sheet.clear();
  sheet.appendRow(["ID", "Сотрудник", "Дата", "ФИО", "Телефон", "Статус", "Заметки"]);
  headerStyle(sheet, 7);
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 180);
  sheet.setFrozenRows(1);
  
  return "✅ Листы созданы: Все лиды, Статистика, Активные лиды";
}

// Стилизация заголовка
function headerStyle(sheet, cols) {
  const range = sheet.getRange(1, 1, 1, cols);
  range.setBackground("#1a1a1a")
       .setFontColor("#00ff88")
       .setFontWeight("bold")
       .setFontSize(12)
       .setVerticalAlignment("middle");
}

// Синхронизация - ГЛАВНАЯ ФУНКЦИЯ
function syncAllSheets() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  
  try {
    const allLeads = getAllLeads();
    
    // 1. Все лиды - ЦЕНТРАЛЬНЫЙ ЛИСТ
    const leadsSheet = ss.getSheetByName(LEADS_SHEET);
    if (!leadsSheet) return "❌ Не найден лист 'Все лиды'";
    
    leadsSheet.clearContents();
    leadsSheet.appendRow(["ID", "Сотрудник", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"]);
    headerStyle(leadsSheet, 10);
    
    // Сортировка по дате (новее сверху)
    allLeads.sort((a, b) => {
      const dA = a.date ? new Date(a.date).getTime() : 0;
      const dB = b.date ? new Date(b.date).getTime() : 0;
      return dB - dA;
    });
    
    if (allLeads.length > 0) {
      const data = allLeads.map((l, i) => {
        // защита от null/undefined
        return [
          i + 1,
          l.name || "",
          l.date || "",
          l.vacancy || "",
          l.city || "",
          l.fio || "",
          l.phone || "",
          l.age || "",
          l.status || "",
          l.notes || ""
        ];
      });
      
      const range = leadsSheet.getRange(2, 1, data.length, 10);
      range.setValues(data);
      
      // Чередующаяся заливка (тёмная тема)
      for (let i = 0; i < data.length; i++) {
        leadsSheet.getRange(2 + i, 1, 1, 10).setBackground(i % 2 === 0 ? "#2d2d2d" : "#1a1a1a");
      }
    }
    
    // 2. Статистика
    const statsSheet = ss.getSheetByName(STATS_SHEET);
    if (statsSheet) {
      statsSheet.clearContents();
      statsSheet.appendRow(["Дата", "Всего", "Подписано", "Отказ", "Свяьь", "Активные", "Успех %"]);
      headerStyle(statsSheet, 7);
      
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
      
      // границы
      statsSheet.getRange(1, 1, 2, 7).setBorder(true, true, true, true, true, true);
    }
    
    // 3. Активные лиды
    const activesSheet = ss.getSheetByName(ACTIVES_SHEET);
    if (activesSheet) {
      activesSheet.clearContents();
      activesSheet.appendRow(["ID", "Сотрудник", "Дата", "ФИО", "Телефон", "Статус", "Заметки"]);
      headerStyle(activesSheet, 7);
      
      const actives = allLeads.filter(l => ACTIVE_STATUSES.indexOf(l.status || "") >= 0);
      if (actives.length > 0) {
        const data = actives.map((l, i) => {
          return [
            i + 1,
            l.name || "",
            l.date || "",
            l.fio || "",
            l.phone || "",
            l.status || "",
            l.notes || ""
          ];
        });
        
        const range = activesSheet.getRange(2, 1, data.length, 7);
        range.setValues(data);
        
        // Чередующаяся заливка
        for (let i = 0; i < data.length; i++) {
          activesSheet.getRange(2 + i, 1, 1, 7).setBackground(i % 2 === 0 ? "#2d2d2d" : "#1a1a1a");
        }
        
        // границы
        range.setBorder(true, true, true, true, true, true);
      }
    }
    
    return "✅ Синхронизация завершена:\n✅ Сотрудников: " + EMPLOYEES.length + "\n✅ Всего лидов: " + allLeads.length;
  } catch (e) {
    Logger.log("❗ Ошибка syncAllSheets: " + e);
    return "❌ Ошибка: " + e;
  }
}

// Меню в Google Sheets
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("_Офис Молодость Admin v3.1")
      .addItem("Создать листы", "initSystem")
      .addItem("Синхронизироватьall", "syncAllSheets")
      .addToUi();
  } catch (e) {
    Logger.log("Меню не создано: " + e);
  }
}
