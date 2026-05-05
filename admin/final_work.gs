/**
 * ADMIN CRM — Office Молодость v4.0
 * Полностью рабочая версия с дизайном и статистикой
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

function safeReadSheet(emp) {
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
      rows.push({name: emp.name, date: r[0]||"", vacancy: r[1]||"", city: r[2]||"", fio: r[3]||"", phone: r[4]||"", age: r[5]||"", status: r[6]||"", notes: r[7]||""});
    }
    return rows;
  } catch(e) { Logger.log("Ошибка "+emp.name+": "+e); return []; }
}

function getAllLeads() {
  let all = [];
  EMPLOYEES.forEach(emp => all = all.concat(safeReadSheet(emp)));
  Logger.log("ВСЕГО ЛИДОВ: "+all.length);
  return all;
}

function initSystem() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  let s = ss.getSheetByName(LEADS_SHEET) || ss.insertSheet(LEADS_SHEET); s.clear(); s.appendRow(["ID","Сотрудник","Дата","Вакансия","Город","ФИО","Телефон","Возраст","Статус","Заметки"]); headerStyle(s,10);
  s = ss.getSheetByName(STATS_SHEET) || ss.insertSheet(STATS_SHEET); s.clear(); s.appendRow(["Дата","Всего","Подписано","Отказ","Свяьь","Активные","Успех %"]); headerStyle(s,7);
  s = ss.getSheetByName(ACTIVES_SHEET) || ss.insertSheet(ACTIVES_SHEET); s.clear(); s.appendRow(["ID","Сотрудник","Дата","ФИО","Телефон","Статус","Заметки"]); headerStyle(s,7);
  return "✅ Листы созданы";
}

function headerStyle(sheet,cols) {
  const r = sheet.getRange(1,1,1,cols); r.setBackground("#1a1a1a").setFontColor("#00ff88").setFontWeight("bold").setFontSize(12);
}

function syncAllSheets() {
  const ss = SpreadsheetApp.openById(ADMIN_SPREADSHEET_ID);
  const allLeads = getAllLeads();
  const ls = ss.getSheetByName(LEADS_SHEET); if(!ls) return "❌ Лист не найден";
  ls.clearContents(); ls.appendRow(["ID","Сотрудник","Дата","Вакансия","Город","ФИО","Телефон","Возраст","Статус","Заметки"]); headerStyle(ls,10);
  allLeads.sort((a,b)=>{const dA=a.date?new Date(a.date).getTime():0;const dB=b.date?new Date(b.date).getTime():0;return dB-dA;});
  if(allLeads.length>0){
    const data = allLeads.map((l,i)=>[i+1,l.name||"",l.date||"",l.vacancy||"",l.city||"",l.fio||"",l.phone||"",l.age||"",l.status||"",l.notes||""]);
    const rng = ls.getRange(2,1,data.length,10); rng.setValues(data);
    for(let i=0;i<data.length;i++) ls.getRange(2+i,1,1,10).setBackground(i%2===0?"#2d2d2d":"#1a1a1a");
  }
  const st = ss.getSheetByName(STATS_SHEET); if(st){st.clearContents();st.appendRow(["Дата","Всего","Подписано","Отказ","Свяьь","Активные","Успех %"]);headerStyle(st,7);let tot=allLeads.length,conn=0,rfs=0,ct=0,ac=0;allLeads.forEach(l=>{const s=l.status||"";if(s.includes("✅"))conn++;if(s.includes("🔴")||s.includes("⚫")||s.includes("❌"))rfs++;if(s.includes("💬")||s.includes("🟡"))ct++;if(ACTIVE_STATUSES.indexOf(s)>=0)ac++;});st.appendRow([new Date(),tot,conn,rfs,ct,ac,(tot>0?Math.round(100*conn/tot):0)+"%"]);}
  const acs = ss.getSheetByName(ACTIVES_SHEET); if(acs){acs.clearContents();acs.appendRow(["ID","Сотрудник","Дата","ФИО","Телефон","Статус","Заметки"]);headerStyle(acs,7);const act=allLeads.filter(l=>ACTIVE_STATUSES.indexOf(l.status||"")>=0);if(act.length>0){const data=act.map((l,i)=>[i+1,l.name||"",l.date||"",l.fio||"",l.phone||"",l.status||"",l.notes||""]);acs.getRange(2,1,data.length,7).setValues(data);}}
  return "✅ Синхронизация: "+allLeads.length+" лидов";
}

function onOpen(){try{const ui=SpreadsheetApp.getUi();ui.createMenu("_Офис Молодость Admin v4.0").addItem("Создать листы","initSystem").addItem("Синхронизировать","syncAllSheets").addToUi();}catch(e){Logger.log("Меню: "+e);}}
