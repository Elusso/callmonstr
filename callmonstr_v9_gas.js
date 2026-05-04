/**
 * CallMonstr v9.0 — Google Apps Script
 * Офис Молодость | Темная тема | Полный контроль
 * 
 * Features:
 * - Dark theme (#1a1a1a, #00ff88, Comfortaa)
 * - Auto-create sheets with styling
 * - Transfer leads between employees (NO modification of existing sheets)
 * - Sparklines in stats
 * - "Отработать" sheet in each employee's table
 * 
 * Usage: Alt+F11 → Paste → Save → Run
 */

// ─── GLOBAL CONFIG ─────────────────────────────────────────────────────────────
const ADMIN_SPREADSHEET_ID = "1AUCWikqIhAGxXKVvTm1SLJjmLoUh0Y_c7Sry8X8cXos";
const SHUFFLE_DAYS = 3;
const ARCHIVE_DAYS = 7;

const THEME = {
  bgDeep: "#1a1a1a",
  bgMid: "#2d2d2d",
  accent: "#00ff88",
  text: "#e0e0e0",
  textHeader: "#ffffff",
  font: "Comfortaa"
};

const EMPLOYEES = [
  {name: "Тёмыч",   id: "1VCVAZhTl4cv9T1J4AyzknYqekMy6ZEymlQRkvOB4yJE"},
  {name: "Влад",    id: "1Lt9BmIVShNFserfYacxII6WjGPDn5ObNeiOo9z63oI8"},
  {name: "Соня",    id: "1U2uq6xhVXxcUvN3bTwj0eleQZ5mUm4c4OCpwN-deqoA"},
  {name: "Костян",  id: "1CcaWPBvdPZ5WwegxDKOQhxljvMHwPZnYLgpl05QVB1c"},
  {name: "Денишк",  id: "1pYRyigxMNSmrqr92RZ9-I5rbehErgX-PSiRbhn4Z4qU"}
];

const MAIN_HEADERS = ["Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"];
const ACTIVE_HEADERS = ["Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Статус", "Выезд", "Приезд", "Заметки"];
const STATS_HEADERS = ["Период", "Сотрудник", "Всего", "Дозвон", "НД", "Подписано", "Комиссовано", "%", "Возраст", "График"];
const HISTORY_HEADERS = ["Дата", "Сотрудник", "ФИО", "Телефон", "Статус (Старый)", "Статус (Новый)", "Кто изменил"];
const WORKED_COL_HEADERS = ["Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Статус", "Заметки"];

const VALID_STATUSES = [
  "⚪️ Новый", "🔴 НД", "🤔 ДУМ", "🟡 ПЕРЕЗВОНИТЬ", "💬 СВЯЗЬ МЕССЕНДЖЕР",
  "✅ ПОДПИСАН", "✅ Подписан", "⚫️ ОТКАЗ", "❌ Отказ", "📝 Заявка",
  "🎫 Ожидает билеты", "🚗 Ожидает выезда", "🚀 В пути", "🏛 В военкомате",
  "🔍 На проверке", "🎗 Комиссован"
];

// ─── UTILITY FUNCTIONS ───────────────────────────────────────────────────────

function isValidStatus(s) { return VALID_STATUSES.includes((s||"").trim()); }
function isActiveStatus(s) { if(!s) return false; s=s.trim().toLowerCase(); return s.includes("подписан")||s.includes("пути")||s.includes("военкомат")||s.includes("проверке")||s.includes("заявка")||s.includes("билеты")||s.includes("выезда")||s.includes("комиссован"); }
function isND(s) { return (s||"").toLowerCase().includes("нд"); }
function isPodpisan(s) { return (s||"").toLowerCase().includes("подписан"); }
function isKomissovan(s) { return (s||"").toLowerCase().includes("комиссован"); }

function sparkline(values, width=10) {
  if(!values||values.length==0) return "▁".repeat(width);
  const min=Math.min(...values), max=Math.max(...values);
  if(min==max) return "▄".repeat(width);
  const bars="▁▂▃▄▅▆▇█";
  let r=""; for(const v of values){const i=Math.floor(((v-min)/(max-min))*(bars.length-1));r+=bars[i];} 
  return r.slice(0,width);
}

// ─── INITIALIZATION ──────────────────────────────────────────────────────────

function initFullSystem() {
  const ss=SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  
  // Create admin sheets with dark theme
  const sheetsConfig=[
    ["📒 Инструкция", [], 4, true],
    ["Все лиды", MAIN_HEADERS, MAIN_HEADERS.length, false],
    ["🎯 Активные", ACTIVE_HEADERS, ACTIVE_HEADERS.length, false],
    ["👥 Сотрудники", ["Сотрудник", "ID", "Ссылка", "Статус", "Дата", "Лидов", "Дозвон %"], 7, false],
    ["📊 Статистика", STATS_HEADERS, STATS_HEADERS.length, false],
    ["📝 История", HISTORY_HEADERS, HISTORY_HEADERS.length, false],
    ["🗄 Архив", ["Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Финал"], 7, false],
    ["📝 Лог", ["Дата", "Сотрудник", "ФИО", "Ошибка"], 4, false],
  ];
  
  const existing={}; ss.getSheets().forEach(s=>existing[s.getSheetName()]=s);
  sheetsConfig.forEach(([name, headers, cols, isInst]) => {
    const ws=existing[name]?existing[name].clear() && existing[name]:ss.insertSheet(name);
    if(!isInst && headers.length>0) {
      ws.appendRow(headers);
      const r=ws.getRange(1,1,1,headers.length);
      r.setBackground(THEME.bgDeep).setFontColor(THEME.accent).setFontFamily(THEME.font).setFontWeight("bold").setFontSize(11);
      ws.setFrozenRows(1); ws.setRowHeight(1, 40);
      if(name.includes("лиды")||name.includes("Активные")) {
        [120,180,140,220,140,80,180,300,150,150].forEach((w,i)=>{try{ws.setColumnWidth(i+1,w);}catch{}});
      }
    }
  });
  
  // Setup instruction sheet
  const inst=ss.getSheetByName("📒 Инструкция");
  if(inst) {
    inst.clear();
    inst.getRange("B1:C1").setValue("CALLMONSTR v9.0").merge();
    inst.getRange("B2:C2").setValue("ТЕМНЫЙ ТЕРМИНАЛ АДМИНА").merge();
    inst.getRange("B3:C3").setValue("Офис Молодость").merge().setFontSize(20).setFontWeight("bold");
    inst.getRange("B3:C3").setBackgroundColor(THEME.accent);
    
    const blocks=[
      ["📞 УПРАВЛЕНИЕ ЛИДАМИ", "sync() → Синхронизация\nshuffle() → Шаттл НД (>3 дней)"],
      ["📊 АНАЛИТИКА", "stats() → Статистика\nhist() → История смен статусов"],
      ["🤖 АВТОМАТИЗАЦИЯ", "archive() → Архив старых НД\nbackup() → Бэкап в Drive"],
    ];
    let row=4;
    blocks.forEach(([t,d])=>{inst.getRange("B"+row).setValue(t); inst.getRange("C"+(row+1)+":D"+(row+1)).setValue(d).merge(); inst.setRowHeight(row,30); inst.setRowHeight(row+1,70); row+=5;});
  }
  
  // Setup employee sheets - CREATE "Отработать" sheet ONLY
  EMPLOYEES.forEach(emp => {
    try {
      const ess=SpreadsheetApp.openById(emp.id);
      let ws=ess.getSheetByName("Отработать");
      if(!ws) ws=ess.insertSheet("Отработать");
      ws.clear();
      ws.appendRow(WORKED_COL_HEADERS);
      const r=ws.getRange(1,1,1,WORKED_COL_HEADERS.length);
      r.setBackground(THEME.accent).setFontColor("#000").setFontFamily(THEME.font).setFontWeight("bold").setFontSize(11);
      ws.setFrozenRows(1);
      [120,180,140,220,140,80,180,300].forEach((w,i)=>{try{ws.setColumnWidth(i+1,w);}catch{}});
      ws.setRowHeight(1,40);
    } catch(e) {
      Logger.log("⚠️ Не могу создать Отработать для "+emp.name+": "+e.message);
    }
  });
  
  // Setup employee sheet in admin
  const empSheet=ss.getSheetByName("👥 Сотрудники");
  if(empSheet) {
    EMPLOYEES.forEach(e => {
      const row=empSheet.getRange("A:A").getValues().flat().indexOf(e.name);
      if(row<0) {
        empSheet.appendRow([e.name, e.id, "https://docs.google.com/spreadsheets/d/"+e.id, "", "", 0, "0%"]);
      } else {
        empSheet.getRange(row+1, 3).setValue("https://docs.google.com/spreadsheets/d/"+e.id);
      }
    });
  }
  
  return "✅ Все листы созданы. 'Отработать' добавлен в каждую таблицу сотрудника.";
}

// ─── DATA SYNC ───────────────────────────────────────────────────────────────

function syncAllData() {
  const ss=SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const leads=ss.getSheetByName("Все лиды"), active=ss.getSheetByName("🎯 Активные");
  const hist=ss.getSheetByName("📝 История"), logSheet=ss.getSheetByName("📝 Лог");
  if(!leads||!active||!hist||!logSheet) return "❌ Не найдены листы. Выполните initFullSystem()";
  
  if(leads.getLastRow()>1) leads.deleteRows(2, leads.getLastRow()-1);
  if(active.getLastRow()>1) active.deleteRows(2, active.getLastRow()-1);
  
  const ldata=[], adata=[], hdata=[], errs=[];
  
  EMPLOYEES.forEach(emp => {
    try {
      const ess=SpreadsheetApp.openById(emp.id);
      // Try "Лиды" first, then sheet1
      let ws=ess.getSheetByName("Лиды");
      if(!ws) ws=ess.getSheets()[0];
      if(!ws || ws.getLastRow()<2) return;
      
      const rows=ws.getDataRange().getValues();
      rows.slice(1).forEach(row => {
        if(row.length<7) return;
        const fio=(row[3]||"").toString().trim(), phone=(row[4]||"").toString().trim(), status=(row[6]||"").toString().trim();
        if(!fio && !phone) return;
        if(status && !isValidStatus(status)) { errs.push([new Date(), emp.name, fio, "Невалидный: "+status]); return; }
        ldata.push([row[0]||new Date(), row[1]||"", row[2]||"", fio, phone, row[5]||"", status, row[7]||"", emp.name, new Date()]);
        if(isActiveStatus(status)) adata.push([row[0]||new Date(), emp.name, fio, phone, row[1]||"", row[2]||"", status, "", "", row[7]||""]);
        hdata.push([new Date(), emp.name, fio, phone, "-", status, "System Sync"]);
      });
    } catch(e) { errs.push([new Date(), emp.name, "-", "Ошибка: "+e.message]); }
  });
  
  if(ldata.length) leads.getRange(leads.getLastRow()+1,1,ldata.length,ldata[0].length).setValues(ldata);
  if(adata.length) active.getRange(active.getLastRow()+1,1,adata.length,adata[0].length).setValues(adata);
  if(hdata.length) hist.getRange(hist.getLastRow()+1,1,hdata.length,hdata[0].length).setValues(hdata);
  if(errs.length) logSheet.getRange(logSheet.getLastRow()+1,1,errs.length,errs[0].length).setValues(errs);
  
  return "✅ Синхронизация: "+ldata.length+" лидов";
}

// ─── SHUFFLE (TRANSFER) ─────────────────────────────────────────────────────

function shuffleOldND() {
  const ss=SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Все лиды");
  if(!sheet) return "❌ Лист 'Все лиды' не найден";
  
  const data=sheet.getDataRange().getValues();
  if(data.length<=1) return "⏸ Нет лидов";
  
  // Count leads per employee
  const counts={};
  EMPLOYEES.forEach(e=>counts[e.name]=0);
  data.slice(1).forEach(r=>{if(r.length>8 && counts[r[8]]!==undefined) counts[r[8]]++});
  
  // Find employee with least leads
  let target=EMPLOYEES[0], min=Infinity;
  EMPLOYEES.forEach(e=>{if(counts[e.name]<min){min=counts[e.name];target=e;}});
  
  const today=new Date();
  let moved=0;
  
  // Transfer old ND leads to target employee
  for(let i=data.length-1; i>=1; i--) {
    const r=data[i];
    if(r.length<9 || r[6]!=="🔴 НД") continue;
    
    const d=parseDate(r[0]);
    if(!d) continue;
    const daysOld=(today-d)/(1000*60*60*24);
    
    if(daysOld>SHUFFLE_DAYS && r[8]!==target.name) {
      // Update employee in "Все лиды"
      sheet.getRange(i+1, 9).setValue(target.name);
      moved++;
      
      // Also add to target's "Отработать" sheet
      try {
        const targetSS=SpreadsheetApp.openById(target.id);
        const workSheet=targetSS.getSheetByName("Отработать");
        if(workSheet) {
          workSheet.appendRow([r[0]||today, target.name, r[3]||"", r[4]||"", r[1]||"", r[2]||"", r[6]||"", r[7]||""]);
        }
      } catch(e) {
        Logger.log("⚠️ Не могу добавить в Отработать для "+target.name);
      }
    }
  }
  
  return "🚀 Шаттл завершен! Перенесено: "+moved+" лидов";
}

// ─── STATISTICS ──────────────────────────────────────────────────────────────

function updateStats() {
  const ss=SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet=ss.getSheetByName("📊 Статистика");
  if(!sheet) return "❌ Лист '📊 Статистика' не найден";
  
  if(sheet.getLastRow()>1) sheet.deleteRows(2, sheet.getLastRow()-1);
  
  const rows=[];
  EMPLOYEES.forEach(emp => {
    let total=0, nd=0, podp=0, kom=0, ages=[];
    const leads=ss.getSheetByName("Все лиды");
    if(leads) {
      const data=leads.getDataRange().getValues();
      for(let i=1; i<data.length; i++) {
        const r=data[i];
        if(r.length<9 || r[8]!==emp.name) continue;
        total++; if(isND(r[6])) nd++; if(isPodpisan(r[6])) podp++; if(isKomissovan(r[6])) kom++;
        try { const a=parseInt(r[5]); if(a>0) ages.push(a); } catch{}
      }
    }
    if(total>0) {
      const pct=((total-nd)/total)*100;
      const spark=sparkline([total,total-nd,nd,podp], 8);
      rows.push(["Сейчас", emp.name, total, total-nd, nd, podp, kom, pct.toFixed(1)+"%", ages.length?Math.round(ages.reduce((a,b)=>a+b,0)/ages.length):0, spark]);
    }
  });
  
  if(rows.length) {
    sheet.getRange(sheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
    // Color coding
    for(let i=0;i<rows.length;i++) {
      const pct=parseFloat(rows[i][7].replace("%",""));
      const cell=sheet.getRange(i+2,8);
      if(pct>70) cell.setBackground("#39ff14");
      else if(pct>40) cell.setBackground("#ffa502");
      else cell.setBackground("#ff3f34");
    }
  }
  return "✅ Статистика: "+rows.length+" сотрудников";
}

// ─── EXISTING FUNCTIONS ──────────────────────────────────────────────────────

function highlightDuplicates() {
  const ss=SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Все лиды");
  if(!sheet) return "❌ Лист не найден";
  const data=sheet.getDataRange().getValues(), seen={}, dups=[];
  for(let i=1; i<data.length; i++) {
    const r=data[i];
    if(r.length<5) continue;
    const fio=(r[3]||"").toString().trim().toLowerCase(), phone=(r[4]||"").toString().replace(/\D/g,"");
    if(fio && phone) { const k=fio+"_"+phone; if(seen[k]) dups.push([seen[k], i+1]); else seen[k]=i+1; }
  }
  return dups.length ? "🔍 Дубликатов: "+dups.length : "✅ Дубликатов нет";
}

function backupAllTables() {
  let folder=null;
  try { folder=DriveApp.getFoldersByName("CallMonstr_Backups").next(); } catch{}
  if(!folder) folder=DriveApp.createFolder("CallMonstr_Backups");
  
  const log=["💾 Бэкап:"];
  EMPLOYEES.forEach(emp => {
    try {
      const ess=SpreadsheetApp.openById(emp.id);
      const cpy=ess.makeCopy("Backup_"+emp.name+"_"+new Date().toISOString());
      folder.addFile(DriveApp.getFileById(cpy.getId()));
      log.push("✅ "+emp.name);
    } catch(e) { log.push("❌ "+emp.name+": "+e.message); }
  });
  return log.join("\n");
}

function showMenu() {
  return [
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "   Офис Молодость — CallMonstr v9.0",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "init()      — Создать листы и 'Отработать'",
    "sync()      — Синхронизация лидов",
    "shuffle()   — Шаттл НД (>3 дней)",
    "stats()     — Статистика с графиками",
    "backup()    — Бэкап в Google Drive",
    "dupes()     — Поиск дубликатов",
    "full()      — Полный цикл (init+sync+backup)",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
  ].join("\n");
}

function fullCycle() {
  let res="🚀 ПОЛНЫЙ ЦИКЛ\n\n--- Инициализация ---\n";
  res+=initFullSystem()+"\n";
  res+="\n--- Синхронизация ---\n"+syncAllData()+"\n";
  res+="\n--- Бэкап ---\n"+backupAllTables()+"\n";
  res+="\n✅ ПОЛНЫЙ ЦИКЛ ЗАВЕРШЕН";
  return res;
}

function onOpen() {
  const ui=SpreadsheetApp.getUi();
  ui.createMenu("_CALLMONSTR v9.0")
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

function main(args=[]) {
  if(!args || args.length===0) return showMenu();
  const cmd=(args[0]||"").toLowerCase();
  switch(cmd) {
    case "init": return initFullSystem();
    case "sync": return syncAllData();
    case "shuffle": return shuffleOldND();
    case "stats": return updateStats();
    case "backup": return backupAllTables();
    case "dupes": return highlightDuplicates();
    case "full": return fullCycle();
    default: return showMenu();
  }
}

// ─── TRIGGERS ────────────────────────────────────────────────────────────────

function setupTriggers() {
  // Delete existing triggers
  ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t));
  
  // Daily sync at 9:00
  ScriptApp.newTrigger("syncAllData").timeBased().atHour(9).everyDays(1).create();
  
  // Daily stats update at 10:00
  ScriptApp.newTrigger("updateStats").timeBased().atHour(10).everyDays(1).create();
  
  // Daily archive at 23:00
  ScriptApp.newTrigger("archiveOldND").timeBased().atHour(23).everyDays(1).create();
  
  // Weekly backup on Sunday at 22:00
  ScriptApp.newTrigger("backupAllTables").timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(22).create();
}

function archiveOldND() {
  const ss=SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Все лиды");
  if(!sheet || sheet.getLastRow()<=1) return "⏸ Нет данных";
  
  const today=new Date();
  const archivedRows=[], toDelete=[];
  
  const data=sheet.getDataRange().getValues();
  for(let i=data.length-1; i>=1; i--) {
    const r=data[i];
    if(r.length<7) continue;
    const status=r[6]||"";
    if(!isND(status)) continue; // Only archive ND and old leads
    
    const d=parseDate(r[0]);
    if(!d) continue;
    const daysOld=(today-d)/(1000*60*60*24);
    
    if(daysOld>ARCHIVE_DAYS) {
      // Archive to "🗄 Архив"
      const arch=ss.getSheetByName("🗄 Архив");
      if(arch) arch.appendRow([r[0], r[1], r[2], r[3], r[4], r[5], "архивирован "+new Date().toISOString().split("T")[0]]);
      toDelete.push(i+1);
    }
  }
  
  // Delete archived rows (in reverse order)
  toDelete.sort((a,b)=>b-a).forEach(row=>sheet.deleteRows(row));
  
  return "🏗 Архивация завершена! Удалено: "+toDelete.length+" записей";
}

function parseDate(d) {
  if(!d) return null;
  if(d instanceof Date) return d;
  try { return new Date(d); } catch { return null; }
}
