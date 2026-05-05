/**
 * ADMIN CRM — Офис Молодость (v2.3)
 * =================================
 * Минималистичная версия: только сбор данных, без оформления
 * Считывает лиды со всех сотрудников по фиксированной структуре:
 *   A=Дата, B=Вакансия, C=Город, D=ФИО, E=Телефон, F=Возраст, G=Статус, H=Заметки
 * 
 * Листы:
 * - Все лиды: все лиды собранные
 * - Статистика: сводные показатели
 * - Активные лиды: только "в активной работе"
 */

const ADMIN_SPREADSHEET_ID = "1AUCWikqIhAGxXKVvTm1SLJjmLoUh0Y_c7Sry8X8cXos";

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

// Статусы "в активной работе"
const ACTIVE_STATUSES = ["📝 Заявка", "🎫 Ожидает билеты", "🚗 Ожидает выезда", "🚀 В пути", "🏛️ В военкомате", "🔍 На проверке", "✅ ПОДПИСАН", "✅ Подписан"];

// Считывание лидов
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
    } catch (e) {
      Logger.log(`⚠️ ${emp.name}: ${e}`);
    }
  });
  return all;
}

// Создание листов (минимум)
function initSystem() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  
  // "Все лиды"
  let sheet = ss.getSheetByName(LEADS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LEADS_SHEET);
    sheet.appendRow(["ID", "Сотрудник", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"]);
  }
  
  // "Статистика"
  sheet = ss.getSheetByName(STATS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(STATS_SHEET);
    sheet.appendRow(["Дата", "Всего", "Подписано", "Отказ", "Свяьь", "Активные", "Успех %"]);
  }
  
  // "Активные лиды"
  sheet = ss.getSheetByName(ACTIVES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ACTIVES_SHEET);
    sheet.appendRow(["ID", "Сотрудник", "Дата", "ФИО", "Телефон", "Статус", "Заметки"]);
  }
  
  return "✅ Листы созданы";
}

// Синхронизация
function syncAllSheets() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const allLeads = getAllLeads();
  
  // 1. "Все лиды"
  let sheet = ss.getSheetByName(LEADS_SHEET);
  if (!sheet) return "❌ Не найден лист 'Все лиды'";
  
  // очистка
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  
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
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
  
  // 2. "Статистика"
  sheet = ss.getSheetByName(STATS_SHEET);
  if (sheet) {
    sheet.clear();
    sheet.appendRow(["Дата", "Всего", "Подписано", "Отказ", "Свяьь", "Активные", "Успех %"]);
    
    let total = allLeads.length;
    let connected = 0, refused = 0, contact = 0, activeCount = 0;
    
    allLeads.forEach(l => {
      if (l.status.includes("✅")) connected++;
      if (l.status.includes("🔴") || l.status.includes("⚫") || l.status.includes("❌")) refused++;
      if (l.status.includes("💬") || l.status.includes("🟡")) contact++;
      if (ACTIVE_STATUSES.indexOf(l.status) >= 0) activeCount++;
    });
    
    const successRate = total > 0 ? Math.round(100 * connected / total) : 0;
    sheet.appendRow([new Date(), total, connected, refused, contact, activeCount, successRate + "%"]);
  }
  
  // 3. "Активные лиды"
  sheet = ss.getSheetByName(ACTIVES_SHEET);
  if (sheet) {
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    
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
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  }
  
  return "✅ Синхронизация: " + allLeads.length + " лидов";
}

// Меню
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("_Офис Молодость Admin v2.3")
    .addItem("Создать листы", "initSystem")
    .addItem("Синхронизировать", "syncAllSheets")
    .addToUi();
}
