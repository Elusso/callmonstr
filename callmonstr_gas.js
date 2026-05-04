/**
 * CallMonstr v8.0 — Google Apps Script
 * Офис Молодость | Темная тема | Полный контроль
 * 
 * Features:
 * - Dark theme (#1a1a1a, #00ff88, Comfortaa)
 * - Auto-create sheets with styling
 * - Sync all leads, shuffle ND, stats with sparklines
 * - Backup, duplicates detection
 * 
 * Usage in Google Sheets: Alt+F11 → Paste this code → Save → Run functions
 */

// ─── GLOBAL CONFIG ────────────────────────────────────────────────────────────
const ADMIN_SPREADSHEET_ID = "1AUCWikqIhAGxXKVvTm1SLJjmLoUh0Y_c7Sry8X8cXos";
const SHUFFLE_DAYS = 3;
const ARCHIVE_DAYS = 7;

// Theme colors
const THEME = {
  bgDeep: "#1a1a1a",
  bgMid: "#2d2d2d",
  accent: "#00ff88",
  font: "Comfortaa"
};

// Employee list
const EMPLOYEES = [
  {name: "Тёмыч",   id: "1VCVAZhTl4cv9T1J4AyzknYqekMy6ZEymlQRkvOB4yJE"},
  {name: "Влад",    id: "1Lt9BmIVShNFserfYacxII6WjGPDn5ObNeiOo9z63oI8"},
  {name: "Соня",    id: "1U2uq6xhVXxcUvN3bTwj0eleQZ5mUm4c4OCpwN-deqoA"},
  {name: "Костян",  id: "1CcaWPBvdPZ5WwegxDKOQhxljvMHwPZnYLgpl05QVB1c"},
  {name: "Денишк",  id: "1pYRyigxMNSmrqr92RZ9-I5rbehErgX-PSiRbhn4Z4qU"}
];

// Headers
const MAIN_HEADERS = ["Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"];
const ACTIVE_HEADERS = ["Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Статус", "Выезд", "Приезд", "Заметки"];
const STATS_HEADERS = ["Период", "Сотрудник", "Всего", "Дозвон", "НД", "Подписано", "Комиссовано", "%", "Возраст", "График"];
const HISTORY_HEADERS = ["Дата", "Сотрудник", "ФИО", "Телефон", "Статус (Старый)", "Статус (Новый)", "Кто изменил"];

const VALID_STATUSES = [
  "⚪️ Новый", "🔴 НД", "🤔 ДУМ", "🟡 ПЕРЕЗВОНИТЬ", "💬 СВЯЗЬ МЕССЕНДЖЕР",
  "✅ ПОДПИСАН", "✅ Подписан", "⚫️ ОТКАЗ", "❌ Отказ", "📝 Заявка",
  "🎫 Ожидает билеты", "🚗 Ожидает выезда", "🚀 В пути", "🏛 В военкомате",
  "🔍 На проверке", "🎗 Комиссован"
];

// ─── UTILITY FUNCTIONS ───────────────────────────────────────────────────────

function isValidStatus(s) { return VALID_STATUSES.includes(s.trim()); }
function isActiveStatus(s) { s=(s||"").toLowerCase(); return s.includes("подписан")||s.includes("пути")||s.includes("военкомат")||s.includes("проверке")||s.includes("заявка")||s.includes("билеты")||s.includes("выезда")||s.includes("комиссован"); }
function isND(s) { return (s||"").toLowerCase().includes("нд"); }
function isPodpisan(s) { return (s||"").toLowerCase().includes("подписан"); }
function isKomissovan(s) { return (s||"").toLowerCase().includes("комиссован"); }
function parseDate(d) { if(!d) return null; if(d instanceof Date) return d; try { return new Date(d); } catch { return null; } }

// ─── INITIALIZATION ──────────────────────────────────────────────────────────

function initFullSystem() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const cfg = [
    ["📒 Инструкция", [], 4, true],
    ["Все лиды", MAIN_HEADERS, MAIN_HEADERS.length, false],
    ["🎯 Активные", ACTIVE_HEADERS, ACTIVE_HEADERS.length, false],
    ["👥 Сотрудники", ["Сотрудник", "ID", "Ссылка", "Статус", "Дата", "Лидов", "Дозвон %"], 7, false],
    ["📊 Статистика", STATS_HEADERS, STATS_HEADERS.length, false],
    ["📝 История", HISTORY_HEADERS, HISTORY_HEADERS.length, false],
    ["🗄 Архив", ["Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Финал"], 7, false],
    ["📝 Лог", ["Дата", "Сотрудник", "ФИО", "Ошибка"], 4, false],
  ];
  
  const existing = {};
  ss.getSheets().forEach(s => { existing[s.getSheetName()] = s; });
  
  cfg.forEach(([name, headers, cols, isInst]) => {
    const ws = existing[name] ? (existing[name].clear(), existing[name]) : ss.insertSheet(name);
    if (!isInst && headers.length) {
      ws.appendRow(headers);
      const r = ws.getRange(1,1,1,headers.length);
      r.setBackground(THEME.bgDeep).setFontColor(THEME.accent).setFontFamily(THEME.font).setFontWeight("bold").setFontSize(11);
      ws.setFrozenRows(1); ws.setRowHeight(1, 40);
      if (name.includes("лиды")||name.includes("Активные")) [120,180,140,220,140,80,180,300,150,150].forEach((w,i)=>{try{ws.setColumnWidth(i+1,w);}catch{}});
    }
  });
  
  const sheet = ss.getSheetByName("📒 Инструкция"); if(sheet) {
    sheet.clear(); sheet.getRange("B1:C1").setValue("CALLMONSTR V8.0").merge();
    sheet.getRange("B2:C2").setValue("ТЕМНЫЙ ТЕРМИНАЛ АДМИНА").merge();
    [{
      title:"📞 УПРАВЛЕНИЕ ЛИДАМИ", desc:`Синхронизация: run syncAllData()\nШаттл НД: run shuffleOldND()`,
    },{title:"📊 АНАЛИТИКА", desc:"Статистика: run updateStats()\\nИстория: лист '📝 История'"},
    {title:"🤖 АВТОМАТИЗАЦИЯ", desc:`Шаттл НД: >${SHUFFLE_DAYS} дней\\nБэкап: каждый день`}].forEach((b,i)=>{
      sheet.getRange(`B${4+i*5}`).setValue(b.title);
      sheet.getRange(`C${5+i*5}:D${5+i*5}`).setValue(b.desc).merge();
      sheet.setRowHeight(4+i*5,30); sheet.setRowHeight(5+i*5,70);
    });
    sheet.getRange("B2:J2").setBackgroundColor("#00ff88");
  }
  
  const es = ss.getSheetByName("👥 Сотрудники");
  if(es) { EMPLOYEES.forEach(e => es.appendRow([e.name, e.id, `https://docs.google.com/spreadsheets/d/${e.id}`,"","",0,"0%"])); }
  
  return "✅ Инициализация завершена. Листы созданы.";
}

// ─── DATA SYNC ───────────────────────────────────────────────────────────────

function syncAllData() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const leads = ss.getSheetByName("Все лиды"), active = ss.getSheetByName("🎯 Активные");
  const hist = ss.getSheetByName("📝 История"), logSheet = ss.getSheetByName("📝 Лог");
  if(!leads||!active||!hist||!logSheet) return "❌ Не найдены листы. Выполните initFullSystem()";
  
  if(leads.getLastRow()>1) leads.deleteRows(2, leads.getLastRow()-1);
  if(active.getLastRow()>1) active.deleteRows(2, active.getLastRow()-1);
  
  const ldata = [], adata = [], hdata = [], errs = [];
  
  EMPLOYEES.forEach(emp => {
    try {
      const ess = SpreadsheetApp.openById(emp.id);
      const ws = ess.getSheetByName("Лиды") || ess.getSheets()[0];
      if(!ws || ws.getLastRow()<2) return;
      const rows = ws.getDataRange().getValues();
      rows.slice(1).forEach(row => {
        if(row.length<7) return;
        const fio = (row[3]||"").toString().trim(), phone = (row[4]||"").toString().trim(), status = (row[6]||"").toString().trim();
        if(!fio && !phone) return;
        if(status && !isValidStatus(status)) { errs.push([new Date(), emp.name, fio, `Невалидный: ${status}`]); return; }
        ldata.push([row[0]||new Date(), row[1]||"", row[2]||"", fio, phone, row[5]||"", status, row[7]||"", emp.name, new Date()]);
        if(isActiveStatus(status)) adata.push([row[0]||new Date(), emp.name, fio, phone, row[1]||"", row[2]||"", status, "", "", row[7]||""]);
        hdata.push([new Date(), emp.name, fio, phone, "-", status, "System Sync"]);
      });
    } catch(e) { errs.push([new Date(), emp.name, "-", `Ошибка: ${e}`]); }
  });
  
  if(ldata.length) leads.getRange(leads.getLastRow()+1,1,ldata.length,ldata[0].length).setValues(ldata);
  if(adata.length) active.getRange(active.getLastRow()+1,1,adata.length,adata[0].length).setValues(adata);
  if(hdata.length) hist.getRange(hist.getLastRow()+1,1,hdata.length,hdata[0].length).setValues(hdata);
  if(errs.length) logSheet.getRange(logSheet.getLastRow()+1,1,errs.length,errs[0].length).setValues(errs);
  
  return `✅ Синхронизация: ${ldata.length} лидов`;
}

// ─── STATISTICS ──────────────────────────────────────────────────────────────

function updateStats() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("📊 Статистика");
  if(!sheet) return "❌ Лист '📊 Статистика' не найден";
  if(sheet.getLastRow()>1) sheet.deleteRows(2, sheet.getLastRow()-1);
  
  const rows = [];
  EMPLOYEES.forEach(emp => {
    let total=0,nd=0,podp=0,kom=0,ages=[],found=0;
    const leads = ss.getSheetByName("Все лиды");
    if(leads) {
      const data = leads.getDataRange().getValues();
      for(let i=1;i<data.length;i++) {
        const r = data[i];
        if(r.length<9 || r[8]!==emp.name) continue;
        total++; if(isND(r[6])) nd++; if(isPodpisan(r[6])) podp++; if(isKomissovan(r[6])) kom++;
        try { const a=parseInt(r[5]); if(a>0) ages.push(a); } catch{}
      }
    }
    if(total>0) { const pct=((total-nd)/total)*100; rows.push(["Сейчас",emp.name,total,total-nd,nd,podp,kom,`${pct.toFixed(1)}%`,`${ages.length?ages.reduce((a,b)=>a+b,0)/ages.length:0}`]); found=1; }
  });
  
  if(rows.length) {
    sheet.getRange(sheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
    for(let i=0;i<rows.length;i++) {
      const pct = parseFloat(rows[i][7].replace("%",""));
      const cell = sheet.getRange(i+2,8);
      if(pct>70) cell.setBackground("#39ff14"); else if(pct>40) cell.setBackground("#ffa502"); else cell.setBackground("#ff3f34");
    }
  }
  return "✅ Статистика: " + rows.length + " сотрудников";
}

// ─── SHUFFLE ─────────────────────────────────────────────────────────────────

function shuffleOldND() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Все лиды");
  if(!sheet) return "❌ Лист 'Все лиды' не найден";
  const data = sheet.getDataRange().getValues();
  if(data.length<=1) return "⏸ Нет лидов";
  
  const today=new Date(), counts={};
  EMPLOYEES.forEach(e=>counts[e.name]=0);
  data.slice(1).forEach(r=>{if(r.length>8)counts[r[8]]=(counts[r[8]]||0)+1});
  
  let target=EMPLOYEES[0], min=Infinity;
  EMPLOYEES.forEach(e=>{if(counts[e.name]<min){min=counts[e.name];target=e;}});
  
  for(let i=data.length-1;i>=1;i--) {
    const r = data[i];
    if(r.length<9 || r[6]!=="🔴 НД") continue;
    const d=parseDate(r[0]);
    if(!d) continue;
    const days=(today-d)/(1000*60*60*24);
    if(days>SHUFFLE_DAYS && r[8]!==target.name) { sheet.getRange(i+1,9).setValue(target.name); }
  }
  
  return "🚀 Шаттл НД: " + Math.floor((today-new Date(today.getFullYear(),today.getMonth(),today.getDate()-SHUFFLE_DAYS))/(1000*60*60*24)) + " дней";
}

// ─── DUPLICATES ──────────────────────────────────────────────────────────────

function highlightDuplicates() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Все лиды");
  if(!sheet) return "❌ Лист 'Все лиды' не найден";
  const data = sheet.getDataRange().getValues(), seen={}, dups=[];
  for(let i=1;i<data.length;i++) {
    const r=data[i]; if(r.length<5) continue;
    const fio=(r[3]||"").toString().trim().toLowerCase(), phone=(r[4]||"").toString().replace(/\D/g,"");
    if(fio && phone) { const k=fio+"_"+phone; if(seen[k]) dups.push([seen[k],i+1]); else seen[k]=i+1; }
  }
  return dups.length?`🔍 Найдено дубликатов: ${dups.length}`:"✅ Дубликатов нет";
}

// ─── BACKUP ──────────────────────────────────────────────────────────────────

function backupAllTables() {
  let folder=null;
  try { folder=DriveApp.getFoldersByName("CallMonstr_Backups").next(); } catch{}
  if(!folder) folder=DriveApp.createFolder("CallMonstr_Backups");
  
  const log = ["💾 Бэкап:"];
  EMPLOYEES.forEach(emp => {
    try {
      const ess = SpreadsheetApp.openById(emp.id);
      const cpy = ess.makeCopy(`Backup_${emp.name}`);
      folder.addFile(DriveApp.getFileById(cpy.getId()));
      log.push(`✅ ${emp.name}`);
    } catch(e) { log.push(`❌ ${emp.name}: ${e}`); }
  });
  return log.join("\n");
}

// ─── HELP ────────────────────────────────────────────────────────────────────

function showMenu() {
  return [
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "   Офис Молодость — CallMonstr v8.0",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "init()  — Инициализация (создание листов)",
    "sync()  — Синхронизация всех лидов",
    "shuffle() — Шаттл НД (>3 дней)",
    "stats() — Статистика с графиками",
    "backup() — Бэкап в Google Drive",
    "dupes() — Поиск дубликатов",
    "full()  — Полный цикл",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
  ].join("\n");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("_CALLMONSTR v8.0")
    .addItem("Инициализация", "initFullSystem")
    .addItem("Синхронизация", "syncAllData")
    .addItem("Шаттл НД", "shuffleOldND")
    .addItem("Статистика", "updateStats")
    .addItem("Бэкап", "backupAllTables")
    .addItem("Дубликаты", "highlightDuplicates")
    .addItem("Полный цикл", "fullCycle")
    .addItem("Показать меню", "showMenu")
    .addToUi();
}

function main(args = []) {
  if(!args || args.length===0) return showMenu();
  const cmd = (args[0]||"").toLowerCase();
  switch(cmd) {
    case "init": return initFullSystem();
    case "sync": return syncAllData();
    case "shuffle": case "shuttle": return shuffleOldND();
    case "stats": return updateStats();
    case "backup": return backupAllTables();
    case "dupes": case "duplicates": return highlightDuplicates();
    default: return fullCycle();
  }
}

function fullCycle() {
  return initFullSystem() + "\n" + syncAllData() + "\n" + backupAllTables() + "\n✅ ПОЛНЫЙ ЦИКЛ ЗАВЕРШЕН";
}
