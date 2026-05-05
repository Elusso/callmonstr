/**
 * ADMIN CRM — Офис Молодость (v2.1)
 * =================================
 * Главная CRM для владельца офиса.
 * Считывает лиды со всех сотрудников по фиксированной структуре:
 *   A=Дата, B=Вакансия, C=Город, D=ФИО, E=Телефон, F=Возраст, G=Статус, H=Заметки
 * 
 * Листы:
 * - Все лиды: все лиды, отсортированные по дате (новее сверху)
 * - Тиеры: сортировка по тиерам (SS/S/A) для удобства анализа
 * - Статистика: сводные показатели по всему офису
 * - Активные лиды: только "в активной работе" (в пути, ожидает выезда и т.д.)
 */

// Global Config
const ADMIN_SPREADSHEET_ID = "1AUCWikqIhAGxXKVvTm1SLJjmLoUh0Y_c7Sry8X8cXos";

const THEME = {
  bgDeep: "#1a1a1a",
  bgMid: "#2d2d2d",
  accent: "#00ff88",
  text: "#e0e0e0",
  textHeader: "#ffffff",
  font: "Comfortaa"
};

// Сотрудники (Влад наравне с другими)
const EMPLOYEES = [
  {name: "Тёмыч",  id: "1VCVAZhTl4cv9T1J4AyzknYqekMy6ZEymlQRkvOB4yJE"},
  {name: "Влад",   id: "1Lt9BmIVShNFserfYacxII6WjGPDn5ObNeiOo9z63oI8"},
  {name: "Соня",   id: "1U2uq6xhVXxcUvN3bTwj0eleQZ5mUm4c4OCpwN-deqoA"},
  {name: "Костян", id: "1CcaWPBvdPZ5WwegxDKOQhxljvMHwPZnYLgpl05QVB1c"},
  {name: "Денишк", id: "1pYRyigxMNSmrqr92RZ9-I5rbehErgX-PSiRbhn4Z4qU"}
];

// Константы листов
const LEADS_SHEET = "Все лиды";
const TIER_SHEET = "Тиеры";
const STATS_SHEET = "Статистика";
const ACTIVES_SHEET = "Активные лиды";

// Статусы → тиер (для аналитики)
const STATUS_TO_TIER = {
  "✅ Подписанные": "SS", "✅ Подписан": "SS", "✅ ПОДПИСАН": "SS", "🎗️ Комиссован": "SS",
  "🔴 НД": "S", "🤔 ДУМ": "S", "⚫ ОТКАЗ": "S", "❌ Отказ": "S",
  "🟡 ПЕРЕЗВОНИТЬ": "A", "💬 СВЯЗЬ МЕССЕНДЖЕР": "A", "🎫 Ожидает билеты": "A",
  "🚗 Ожидает выезда": "A", "🚀 В пути": "A", "🏛️ В военкомате": "A", "🔍 На проверке": "A", "📝 Заявка": "A"
};

function getTier(status) {
  if (!status) return "A";
  return STATUS_TO_TIER[status.trim()] || "A";
}

// Индексы колонок (0-based): A=0, B=1, ..., H=7
const COLS = { date: 0, vacancy: 1, city: 2, fio: 3, phone: 4, age: 5, status: 6, notes: 7 };

// Заголовки
function LEADS_HEADER() { return ["ID", "Сотрудник", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"]; }
function TIER_HEADER() { return ["ID", "Сотрудник", "Дата", "Вакансия", "Город", "ФИО", "Телефон", "Статус", "Тир"]; }
function STATS_HEADER() { return ["Дата", "Всего", "T1 (SS)", "T2 (S)", "T3 (A)", "Подписано", "Отказ", "Свяьь", "Активные", "Успех %"]; }
function ACTIVES_HEADER() { return ["ID", "Сотрудник", "Дата", "ФИО", "Телефон", "Статус", "Заметки"]; }

// Статусы "в активной работе"
const ACTIVE_STATUSES = ["🚀 В пути", "🚗 Ожидает выезда", "🏛️ В военкомате", "🔍 На проверке"];

// Стилизация (темная тема +Comfortaa)
function styleHeader(range) {
  range.setBackground(THEME.bgDeep)
       .setFontColor(THEME.accent)
       .setFontFamily(THEME.font)
       .setFontWeight("bold")
       .setFontSize(12);
}

function setColumnsWidth(sheet, widths) {
  widths.forEach((w, i) => { try { sheet.setColumnWidth(i + 1, w); } catch {} });
  if (sheet.getLastRow() >= 1) sheet.setRowHeight(1, 40);
}

function tierColor(cell, tier) {
  if (tier === "SS") cell.setBackground("#ff3f34");
  else if (tier === "S") cell.setBackground("#ffa502");
  else if (tier === "A") cell.setBackground("#ffff00");
  else cell.setBackground("#ffffff");
}

// Получение данных из таблицы сотрудника
function readEmployeeSheet(emp) {
  try {
    const sheet = SpreadsheetApp.openById(emp.id).getSheets()[0];
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      // Берём только A-H (0-7)
      if (r.length < 8) continue;
      
      const status = r[COLS.status] || "";
      const tier = getTier(status);
      
      rows.push({
        name: emp.name,
        date: r[COLS.date] || "",
        vacancy: r[COLS.vacancy] || "",
        city: r[COLS.city] || "",
        fio: r[COLS.fio] || "",
        phone: r[COLS.phone] || "",
        age: r[COLS.age] || "",
        status: status,
        tier: tier,
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
  }
  
  // "Тиеры"
  if (!sheets.includes(TIER_SHEET)) {
    const sh = ss.insertSheet(TIER_SHEET);
    sh.appendRow(TIER_HEADER());
    styleHeader(sh.getRange(1, 1, 1, TIER_HEADER().length));
    sh.setFrozenRows(1);
    setColumnsWidth(sh, [50, 120, 100, 180, 120, 220, 140, 120, 60]);
  }
  
  // "Статистика"
  if (!sheets.includes(STATS_SHEET)) {
    const sh = ss.insertSheet(STATS_SHEET);
    sh.appendRow(STATS_HEADER());
    styleHeader(sh.getRange(1, 1, 1, STATS_HEADER().length));
    sh.setFrozenRows(1);
    setColumnsWidth(sh, [120, 80, 70, 70, 70, 90, 90, 90, 80, 70]);
  }
  
  // "Активные лиды"
  if (!sheets.includes(ACTIVES_SHEET)) {
    const sh = ss.insertSheet(ACTIVES_SHEET);
    sh.appendRow(ACTIVES_HEADER());
    styleHeader(sh.getRange(1, 1, 1, ACTIVES_HEADER().length));
    sh.setFrozenRows(1);
    setColumnsWidth(sh, [50, 120, 100, 220, 140, 120, 180]);
  }
  
  return "✅ Листы созданы: Все лиды, Тиеры, Статистика, Активные лиды";
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
  
  if (allLeads.length > 0) {
    // Сортировка: новее датой сверху
    allLeads.sort((a, b) => {
      const dA = Date.parse(a.date) || 0;
      const dB = Date.parse(b.date) || 0;
      return dB - dA; // убывание
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
    leadsSheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
  
  // 2. "Тиеры" — сортировка по тиерам (SS → S → A)
  const tierSheet = ss.getSheetByName(TIER_SHEET);
  if (tierSheet) {
    if (tierSheet.getLastRow() > 1) tierSheet.deleteRows(2, tierSheet.getLastRow() - 1);
    
    if (allLeads.length > 0) {
      // Сортировка по тиеру (SS=1, S=2, A=3)
      const sorted = allLeads.slice().sort((a, b) => {
        const rankA = a.tier === "SS" ? 1 : a.tier === "S" ? 2 : 3;
        const rankB = b.tier === "SS" ? 1 : b.tier === "S" ? 2 : 3;
        if (rankA !== rankB) return rankA - rankB;
        // внутри одного тира — по дате (новее сверху)
        const dA = Date.parse(a.date) || 0;
        const dB = Date.parse(b.date) || 0;
        return dB - dA;
      });
      
      const data = sorted.map((l, i) => [
        i + 1,
        l.name,
        l.date,
        l.vacancy,
        l.city,
        l.fio,
        l.phone,
        l.status,
        l.tier
      ]);
      tierSheet.getRange(2, 1, data.length, data[0].length).setValues(data);
      // coloring tiers
      tierSheet.getRange(2, 9, data.length, 1).getCellList().forEach(c => tierColor(c, c.getValue()));
    }
  }
  
  // 3. "Статистика"
  const statsSheet = ss.getSheetByName(STATS_SHEET);
  if (statsSheet) {
    statsSheet.clear();
    statsSheet.appendRow(STATS_HEADER());
    styleHeader(statsSheet.getRange(1, 1, 1, STATS_HEADER().length));
    
    // Подсчёт статистики
    const today = new Date();
    let total = allLeads.length;
    let ssCount = 0, sCount = 0, aCount = 0;
    let connected = 0, refused = 0, contact = 0;
    let activesCount = 0;
    
    for (let i = 0; i < allLeads.length; i++) {
      const st = allLeads[i].status;
      const tier = allLeads[i].tier;
      
      if (tier === "SS") ssCount++;
      else if (tier === "S") sCount++;
      else aCount++;
      
      if (st.includes("✅")) connected++;
      if (st.includes("🔴") || st.includes("⚫") || st.includes("❌")) refused++;
      if (st.includes("💬") || st.includes("🟡")) contact++;
      
      // Активные
      if (ACTIVE_STATUSES.indexOf(st) >= 0) activesCount++;
    }
    
    const successRate = total > 0 ? Math.round(100 * connected / total) : 0;
    
    statsSheet.appendRow([
      today,
      total,
      ssCount,
      sCount,
      aCount,
      connected,
      refused,
      contact,
      activesCount,
      successRate + "%"
    ]);
    
    // границы
    statsSheet.getRange(1, 1, 2, 10).setBorder(true, true, true, true, true, true);
  }
  
  // 4. "Активные лиды" — только "в активной работе"
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
      activesSheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  }
  
  return "✅ Синхронизация: " + allLeads.length + " лидов, " + activesCount(allLeads) + " активных";
}

function activesCount(leads) {
  return leads.filter(l => ACTIVE_STATUSES.indexOf(l.status) >= 0).length;
}

// Сортировка по тиерам (для обновления только листа "Тиеры")
function sortTierList() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TIER_SHEET);
  if (!sheet) return "❌ Лист 'Тиеры' не найден";
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return "⏸ Нет данных";
  
  const rows = data.slice(1);
  
  rows.sort((a, b) => {
    const rankA = a[8] === "SS" ? 1 : a[8] === "S" ? 2 : 3;
    const rankB = b[8] === "SS" ? 1 : b[8] === "S" ? 2 : 3;
    if (rankA !== rankB) return rankA - rankB;
    const dateA = Date.parse(a[2]) || 0;
    const dateB = Date.parse(b[2]) || 0;
    return dateB - dateA;
  });
  
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    sheet.getRange(2, 9, rows.length, 1).getCellList().forEach(c => tierColor(c, c.getValue()));
  }
  
  return "⬆️ Сортировка Тиеров: " + rows.length + " записей";
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
  ui.createMenu("_Офис Молодость Admin v2.1")
    .addItem("Создать листы", "initSystem")
    .addItem("Синхронизировать все таблицы", "syncAllSheets")
    .addItem("Сортировать Тиеры", "sortTierList")
    .addItem("Экспорт JSON", "exportJSON")
    .addToUi();
}
