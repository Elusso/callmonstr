```javascript
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CRM Call Tracking')
    .addItem('Обновить данные', 'refreshData')
    .addItem('Экспортировать в XLSX', 'exportToXLSX')
    .addItem('Создать бэкап', 'createBackup')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function refreshData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const values = range.getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('Статус') + 1;
  const dateCol = headers.indexOf('Дата создания') + 1;
  const lastCallCol = headers.indexOf('Последний звонок') + 1;
  const notesCol = headers.indexOf('Заметки') + 1;
  const sparklineCol = headers.indexOf('График') + 1;
  const sparklineDataCols = [];
  headers.forEach((h, i) => {
    if (h.includes('Звонок') && i !== lastCallCol - 1) sparklineDataCols.push(i + 1);
  });
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  const twoDaysAgo = new Date(today);
  twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
  const threeDaysAgo = new Date(today);
  threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
  const fourDaysAgo = new Date(today);
  fourDaysAgo.setDate(fourDaysAgo.getDate() - 4);
  const fiveDaysAgo = new Date(today);
  fiveDaysAgo.setDate(fiveDaysAgo.getDate() - 5);
  const sixDaysAgo = new Date(today);
  sixDaysAgo.setDate(sixDaysAgo.getDate() - 6);
  const statuses = [
    '⚪ Новый', '🔴 НД', '🤔 ДУМ', '🟡 ПЕРЕЗВОНИТЬ',
    '💬 СВЯЗЬ МЕССЕНДЖЕР', '✅ ПОДПИСАН', '❌ Отказ', '📝 Заявка',
    '🎫 Ожидает билеты', '🚗 Ожидает выезда', '🚀 В пути',
    '🏛️ В военкомате', '🔍 На проверке', '🎗️ Комиссован'
  ];
  const sparklineMap = {
    '⚪ Новый': 'new',
    '🔴 НД': 'noanswer',
    '🤔 ДУМ': 'thinking',
    '🟡 ПЕРЕЗВОНИТЬ': 'callback',
    '💬 СВЯЗЬ МЕССЕНДЖЕР': 'messenger',
    '✅ ПОДПИСАН': 'signed',
    '❌ Отказ': 'refusal',
    '📝 Заявка': 'request',
    '🎫 Ожидает билеты': 'tickets',
    '🚗 Ожидает выезда': 'departure',
    '🚀 В пути': 'enroute',
    '🏛️ В военкомате': 'draft',
    '🔍 На проверке': 'checking',
    '🎗️ Комиссован': 'discharged'
  };
  const sparklineData = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const status = row[statusCol - 1];
    const dateCreated = row[dateCol - 1];
    const lastCall = row[lastCallCol - 1];
    const notes = row[notesCol - 1];
    const sparklineDataPoints = [];
    if (sparklineDataCols.length > 0) {
      sparklineDataCols.forEach(col => {
        const val = row[col - 1];
        if (val instanceof Date) {
          sparklineDataPoints.push(val.getTime());
        } else if (typeof val === 'number') {
          sparklineDataPoints.push(val);
        }
      });
    }
    let sparklineFormula = '';
    if (sparklineDataPoints.length > 0) {
      sparklineFormula = '=SPARKLINE({' + sparklineDataPoints.join(',') + '}, {"charttype","line";"max",100})';
    } else {
      sparklineFormula = '=SPARKLINE({1}, {"charttype","bar";"max",1})';
    }
    sparklineMap[status] && sparklineFormula;
    if (lastCall instanceof Date) {
      const diffDays = Math.floor((today - lastCall) / (1000 * 60 * 60 * 24));
      if (diffDays === 0) sparklineFormula = '=SPARKLINE({1,2,3,4,5}, {"charttype","line";"max",5})';
      else if (diffDays === 1) sparklineFormula = '=SPARKLINE({1,2,3,4}, {"charttype","line";"max",4})';
      else if (diffDays === 2) sparklineFormula = '=SPARKLINE({1,2,3}, {"charttype","line";"max",3})';
      else if (diffDays === 3) sparklineFormula = '=SPARKLINE({1,2}, {"charttype","line";"max",2})';
      else if (diffDays === 4) sparklineFormula = '=SPARKLINE({1}, {"charttype","line";"max",1})';
    }
    if (status === '⚪ Новый' && dateCreated instanceof Date) {
      const daysSince = Math.floor((today - dateCreated) / (1000 * 60 * 60 * 24));
      if (daysSince <= 1) sparklineFormula = '=SPARKLINE({1,2,3,4,5,6,7}, {"charttype","line";"max",7})';
      else if (daysSince <= 3) sparklineFormula = '=SPARKLINE({1,2,3,4,5}, {"charttype","line";"max",5})';
      else if (daysSince <= 7) sparklineFormula = '=SPARKLINE({1,2,3}, {"charttype","line";"max",3})';
      else sparklineFormula = '=SPARKLINE({1}, {"charttype","line";"max",1})';
    }
    if (notes && notes.length > 0) {
      sparklineFormula = '=SPARKLINE({1,2,3,4,5,6,7,8,9,10}, {"charttype","line";"max",10})';
    }
    sparklineData.push([sparklineFormula]);
  }
  sheet.getRange(2, sparklineCol, sparklineData.length, 1).setFormulas(sparklineData);
}

function exportToXLSX() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getDataRange();
  const blob = range.getBlob().setName('CRM_Call_Tracking_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm') + '.xlsx');
  DriveApp.createFile(blob);
  const url = blob.getUrl();
  const html = HtmlService.createHtmlOutput('<script>window.open("' + url + '","_blank");window.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, 'Экспорт завершен');
}

function createBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupSheet = ss.getSheetByName('Бэкап');
  if (!backupSheet) return;
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const logSheet = ss.getSheetByName('Логи');
  if (!logSheet) return;
  const logRange = logSheet.getRange(1, 1, 1, 3);
  logRange.setValues([[timestamp, 'Бэкап создан', 'Успешно']]);
  const backupRow = backupSheet.getLastRow() + 1;
  backupSheet.getRange(backupRow, 1).setValue(timestamp);
  backupSheet.getRange(backupRow, 2).setValue(ss.getUrl());
  backupSheet.getRange(backupRow, 3).setValue('Активный');
}

function createBackupTrigger() {
  ScriptApp.newTrigger('createBackup')
    .timeBased()
    .everyHours(6)
    .create();
}

function createCallLogEntry(action, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Логи');
  if (!logSheet) return;
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const user = Session.getActiveUser().getEmail();
  const row = logSheet.getLastRow() + 1;
  logSheet.getRange(row, 1).setValue(timestamp);
  logSheet.getRange(row, 2).setValue(user);
  logSheet.getRange(row, 3).setValue(action);
  logSheet.getRange(row, 4).setValue(details);
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== 'Звонки') return;
  const row = range.getRow();
  if (row === 1) return;
  const col = range.getColumn();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('Статус') + 1;
  const dateCol = headers.indexOf('Дата создания') + 1;
  const lastCallCol = headers.indexOf('Последний звонок') + 1;
  if (col === statusCol) {
    const newValue = range.getValue();
    if (newValue) {
      const now = new Date();
      sheet.getRange(row, lastCallCol).setValue(now);
      createCallLogEntry('Изменен статус', 'Строка ' + row + ': ' + newValue);
    }
  }
}

function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const statusCol = headers.indexOf('Статус');
  const statusCounts = {};
  for (let i = 1; i < values.length; i++) {
    const status = values[i][statusCol];
    if (status) {
      statusCounts[status] = (statusCounts[status] || 0) + 1;
    }
  }
  const statsSheet = ss.getSheetByName('Статистика');
  if (!statsSheet) {
    ss.insertSheet('Статистика');
  } else {
    statsSheet.clear();
  }
  const stats = statsSheet.getSheetByName('Статистика');
  stats.getRange(1, 1).setValue('Статус');
  stats.getRange(1, 2).setValue('Количество');
  let row = 2;
  for (const [status, count] of Object.entries(statusCounts)) {
    stats.getRange(row, 1).setValue(status);
    stats.getRange(row, 2).setValue(count);
    row++;
  }
  stats.getRange(1, 1, row - 1, 2).setBorder(true, true, true, true, true, true);
}

function createSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Настройки');
  if (!settingsSheet) {
    const sheet = ss.insertSheet('Настройки');
    sheet.getRange(1, 1).setValue('Параметр');
    sheet.getRange(1, 2).setValue('Значение');
    sheet.getRange(2, 1).setValue('Последний бэкап');
    sheet.getRange(2, 2).setValue(new Date().toISOString());
    sheet.getRange(3, 1).setValue('Частота бэкапа (часы)');
    sheet.getRange(3, 2).setValue(6);
    sheet.getRange(4, 1).setValue('Последний экспорт');
    sheet.getRange(4, 2).setValue('');
  }
}

function createSheetsStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = ['Звонки', 'Логи', 'Статистика', 'Бэкап', 'Настройки'];
  const existingSheets = ss.getSheets().map(s => s.getName());
  requiredSheets.forEach(name => {
    if (!existingSheets.includes(name)) {
      ss.insertSheet(name);
    }
  });
  const headers = [
    'ID', 'ФИО', 'Телефон', 'Дата создания', 'Последний звонок', 'Статус',
    'Заметки', 'Комментарий', 'Источник', 'Ответственный', 'График'
  ];
  const sheet = ss.getSheetByName('Звонки');
  if (sheet) {
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    if (lastCol < headers.length) {
      sheet.getRange(1, lastCol + 1, 1, headers.length - lastCol).setValues([headers.slice(lastCol)]);
    }
    if (lastRow < 1000) {
      sheet.getRange(lastRow + 1, 1, 1000 - lastRow, headers.length).setValues(Array(1000 - lastRow).fill(headers));
    }
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(2, 1, 1000, headers.length).setRowHeight(30);
    sheet.getRange(1, 1, 1, headers.length).setRowHeight(40);
  }
}

function createButtons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('CRM Call Tracking');
  menu.addItem('Перезвонить', 'callbackAction')
    .addItem('Добавить заметку', 'addNoteAction')
    .addItem('Сохранить', 'saveAction')
    .addToUi();
}

function callbackAction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getActiveRange();
  const row = range.getRow();
  if (row === 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const phoneCol = headers.indexOf('Телефон') + 1;
  const statusCol = headers.indexOf('Статус') + 1;
  const lastCallCol = headers.indexOf('Последний звонок') + 1;
  const phone = sheet.getRange(row, phoneCol).getValue();
  const status = sheet.getRange(row, statusCol).getValue();
  if (phone) {
    sheet.getRange(row, lastCallCol).setValue(new Date());
    sheet.getRange(row, statusCol).setValue('🟡 ПЕРЕЗВОНИТЬ');
    createCallLogEntry('Перезвонить', 'Телефон: ' + phone);
  }
}

function addNoteAction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getActiveRange();
  const row = range.getRow();
  if (row === 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const notesCol = headers.indexOf('Заметки') + 1;
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Добавить заметку', 'Введите текст заметки:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    const note = response.getResponseText();
    const currentNote = sheet.getRange(row, notesCol).getValue();
    const newNote = currentNote ? currentNote + '\n' + note : note;
    sheet.getRange(row, notesCol).setValue(newNote);
    createCallLogEntry('Добавлена заметка', 'Строка ' + row + ': ' + note);
  }
}

function saveAction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getActiveRange();
  const row = range.getRow();
  if (row === 1) return;
  createCallLogEntry('Сохранено', 'Строка ' + row);
  ui.alert('Данные сохранены');
}

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return HtmlService.createHtmlOutput('<h1>Лист "Звонки" не найден</h1>');
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>CRM Call Tracking</title><style>body{background:#1a1a1a;color:#00ff88;font-family:"Comfortaa",sans-serif;margin:0;padding:20px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #00ff88;padding:8px;text-align:left;height:30px}th{background:#003300;height:40px}button{background:#00ff88;color:#1a1a1a;border:none;padding:6px 12px;margin:4px;cursor:pointer;font-family:"Comfortaa",sans-serif}button:hover{background:#00cc6a}</style><link href="https://fonts.googleapis.com/css2?family=Comfortaa:wght@300;400;700&display=swap" rel="stylesheet"></head><body><h2>CRM Call Tracking</h2><button onclick="window.open(\'' + ss.getUrl() + '\',\'_blank\')">Открыть в Google Sheets</button><table><thead><tr>';
  headers.forEach(h => html += '<th>' + h + '</th>');
  html += '</tr></thead><tbody>';
  for (let i = 1; i < values.length; i++) {
    html += '<tr>';
    values[i].forEach((v, j) => {
      if (j === headers.indexOf('Статус')) {
        html += '<td>' + v + '</td>';
      } else if (j === headers.indexOf('График')) {
        html += '<td><img src="https://chart.googleapis.com/chart?chs=100x30&cht=ls&chd=t:' + (v ? '1,2,3,4,5' : '1') + '&chco=00ff88&chf=bg,s,1a1a1a" alt="График"></td>';
      } else {
        html += '<td>' + (v || '') + '</td>';
      }
    });
    html += '</tr>';
  }
  html += '</tbody></table><script>function refreshData(){window.location.reload()}</script><button onclick="refreshData()">Обновить</button></body></html>';
  return HtmlService.createHtmlOutput(html);
}

function createAll() {
  createSheetsStructure();
  createSettingsSheet();
  createButtons();
  createDashboard();
  createBackupTrigger();
  refreshData();
}

function createDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statsSheet = ss.getSheetByName('Статистика');
  if (!statsSheet) return;
  statsSheet.clear();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const statusCol = headers.indexOf('Статус');
  const statusCounts = {};
  for (let i = 1; i < values.length; i++) {
    const status = values[i][statusCol];
    if (status) {
      statusCounts[status] = (statusCounts[status] || 0) + 1;
    }
  }
  let row = 1;
  for (const [status, count] of Object.entries(statusCounts)) {
    statsSheet.getRange(row, 1).setValue(status);
    statsSheet.getRange(row, 2).setValue(count);
    row++;
  }
  statsSheet.getRange(1, 1, row - 1, 2).setBorder(true, true, true, true, true, true);
  statsSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
}

function exportToXLSX() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getDataRange();
  const blob = range.getBlob().setName('CRM_Call_Tracking_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm') + '.xlsx');
  DriveApp.createFile(blob);
  const url = blob.getUrl();
  const html = HtmlService.createHtmlOutput('<script>window.open("' + url + '","_blank");window.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, 'Экспорт завершен');
}

function createBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupSheet = ss.getSheetByName('Бэкап');
  if (!backupSheet) return;
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const logSheet = ss.getSheetByName('Логи');
  if (!logSheet) return;
  const logRange = logSheet.getRange(1, 1, 1, 3);
  logRange.setValues([[timestamp, 'Бэкап создан', 'Успешно']]);
  const backupRow = backupSheet.getLastRow() + 1;
  backupSheet.getRange(backupRow, 1).setValue(timestamp);
  backupSheet.getRange(backupRow, 2).setValue(ss.getUrl());
  backupSheet.getRange(backupRow, 3).setValue('Активный');
}

function createBackupTrigger() {
  ScriptApp.newTrigger('createBackup')
    .timeBased()
    .everyHours(6)
    .create();
}

function createCallLogEntry(action, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Логи');
  if (!logSheet) return;
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const user = Session.getActiveUser().getEmail();
  const row = logSheet.getLastRow() + 1;
  logSheet.getRange(row, 1).setValue(timestamp);
  logSheet.getRange(row, 2).setValue(user);
  logSheet.getRange(row, 3).setValue(action);
  logSheet.getRange(row, 4).setValue(details);
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== 'Звонки') return;
  const row = range.getRow();
  if (row === 1) return;
  const col = range.getColumn();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('Статус') + 1;
  const dateCol = headers.indexOf('Дата создания') + 1;
  const lastCallCol = headers.indexOf('Последний звонок') + 1;
  if (col === statusCol) {
    const newValue = range.getValue();
    if (newValue) {
      const now = new Date();
      sheet.getRange(row, lastCallCol).setValue(now);
      createCallLogEntry('Изменен статус', 'Строка ' + row + ': ' + newValue);
    }
  }
}

function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const statusCol = headers.indexOf('Статус');
  const statusCounts = {};
  for (let i = 1; i < values.length; i++) {
    const status = values[i][statusCol];
    if (status) {
      statusCounts[status] = (statusCounts[status] || 0) + 1;
    }
  }
  const statsSheet = ss.getSheetByName('Статистика');
  if (!statsSheet) {
    ss.insertSheet('Статистика');
  } else {
    statsSheet.clear();
  }
  const stats = statsSheet.getSheetByName('Статистика');
  stats.getRange(1, 1).setValue('Статус');
  stats.getRange(1, 2).setValue('Количество');
  let row = 2;
  for (const [status, count] of Object.entries(statusCounts)) {
    stats.getRange(row, 1).setValue(status);
    stats.getRange(row, 2).setValue(count);
    row++;
  }
  stats.getRange(1, 1, row - 1, 2).setBorder(true, true, true, true, true, true);
}

function createSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Настройки');
  if (!settingsSheet) {
    const sheet = ss.insertSheet('Настройки');
    sheet.getRange(1, 1).setValue('Параметр');
    sheet.getRange(1, 2).setValue('Значение');
    sheet.getRange(2, 1).setValue('Последний бэкап');
    sheet.getRange(2, 2).setValue(new Date().toISOString());
    sheet.getRange(3, 1).setValue('Частота бэкапа (часы)');
    sheet.getRange(3, 2).setValue(6);
    sheet.getRange(4, 1).setValue('Последний экспорт');
    sheet.getRange(4, 2).setValue('');
  }
}

function createSheetsStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = ['Звонки', 'Логи', 'Статистика', 'Бэкап', 'Настройки'];
  const existingSheets = ss.getSheets().map(s => s.getName());
  requiredSheets.forEach(name => {
    if (!existingSheets.includes(name)) {
      ss.insertSheet(name);
    }
  });
  const headers = [
    'ID', 'ФИО', 'Телефон', 'Дата создания', 'Последний звонок', 'Статус',
    'Заметки', 'Комментарий', 'Источник', 'Ответственный', 'График'
  ];
  const sheet = ss.getSheetByName('Звонки');
  if (sheet) {
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    if (lastCol < headers.length) {
      sheet.getRange(1, lastCol + 1, 1, headers.length - lastCol).setValues([headers.slice(lastCol)]);
    }
    if (lastRow < 1000) {
      sheet.getRange(lastRow + 1, 1, 1000 - lastRow, headers.length).setValues(Array(1000 - lastRow).fill(headers));
    }
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(2, 1, 1000, headers.length).setRowHeight(30);
    sheet.getRange(1, 1, 1, headers.length).setRowHeight(40);
  }
}

function createButtons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('CRM Call Tracking');
  menu.addItem('Перезвонить', 'callbackAction')
    .addItem('Добавить заметку', 'addNoteAction')
    .addItem('Сохранить', 'saveAction')
    .addToUi();
}

function callbackAction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getActiveRange();
  const row = range.getRow();
  if (row === 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const phoneCol = headers.indexOf('Телефон') + 1;
  const statusCol = headers.indexOf('Статус') + 1;
  const lastCallCol = headers.indexOf('Последний звонок') + 1;
  const phone = sheet.getRange(row, phoneCol).getValue();
  const status = sheet.getRange(row, statusCol).getValue();
  if (phone) {
    sheet.getRange(row, lastCallCol).setValue(new Date());
    sheet.getRange(row, statusCol).setValue('🟡 ПЕРЕЗВОНИТЬ');
    createCallLogEntry('Перезвонить', 'Телефон: ' + phone);
  }
}

function addNoteAction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getActiveRange();
  const row = range.getRow();
  if (row === 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const notesCol = headers.indexOf('Заметки') + 1;
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Добавить заметку', 'Введите текст заметки:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    const note = response.getResponseText();
    const currentNote = sheet.getRange(row, notesCol).getValue();
    const newNote = currentNote ? currentNote + '\n' + note : note;
    sheet.getRange(row, notesCol).setValue(newNote);
    createCallLogEntry('Добавлена заметка', 'Строка ' + row + ': ' + note);
  }
}

function saveAction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getActiveRange();
  const row = range.getRow();
  if (row === 1) return;
  createCallLogEntry('Сохранено', 'Строка ' + row);
  ui.alert('Данные сохранены');
}

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return HtmlService.createHtmlOutput('<h1>Лист "Звонки" не найден</h1>');
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>CRM Call Tracking</title><style>body{background:#1a1a1a;color:#00ff88;font-family:"Comfortaa",sans-serif;margin:0;padding:20px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #00ff88;padding:8px;text-align:left;height:30px}th{background:#003300;height:40px}button{background:#00ff88;color:#1a1a1a;border:none;padding:6px 12px;margin:4px;cursor:pointer;font-family:"Comfortaa",sans-serif}button:hover{background:#00cc6a}</style><link href="https://fonts.googleapis.com/css2?family=Comfortaa:wght@300;400;700&display=swap" rel="stylesheet"></head><body><h2>CRM Call Tracking</h2><button onclick="window.open(\'' + ss.getUrl() + '\',\'_blank\')">Открыть в Google Sheets</button><table><thead><tr>';
  headers.forEach(h => html += '<th>' + h + '</th>');
  html += '</tr></thead><tbody>';
  for (let i = 1; i < values.length; i++) {
    html += '<tr>';
    values[i].forEach((v, j) => {
      if (j === headers.indexOf('Статус')) {
        html += '<td>' + v + '</td>';
      } else if (j === headers.indexOf('График')) {
        html += '<td><img src="https://chart.googleapis.com/chart?chs=100x30&cht=ls&chd=t:' + (v ? '1,2,3,4,5' : '1') + '&chco=00ff88&chf=bg,s,1a1a1a" alt="График"></td>';
      } else {
        html += '<td>' + (v || '') + '</td>';
      }
    });
    html += '</tr>';
  }
  html += '</tbody></table><script>function refreshData(){window.location.reload()}</script><button onclick="refreshData()">Обновить</button></body></html>';
  return HtmlService.createHtmlOutput(html);
}

function createDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statsSheet = ss.getSheetByName('Статистика');
  if (!statsSheet) return;
  statsSheet.clear();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const statusCol = headers.indexOf('Статус');
  const statusCounts = {};
  for (let i = 1; i < values.length; i++) {
    const status = values[i][statusCol];
    if (status) {
      statusCounts[status] = (statusCounts[status] || 0) + 1;
    }
  }
  let row = 1;
  for (const [status, count] of Object.entries(statusCounts)) {
    statsSheet.getRange(row, 1).setValue(status);
    statsSheet.getRange(row, 2).setValue(count);
    row++;
  }
  statsSheet.getRange(1, 1, row - 1, 2).setBorder(true, true, true, true, true, true);
  statsSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
}

function exportToXLSX() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getDataRange();
  const blob = range.getBlob().setName('CRM_Call_Tracking_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm') + '.xlsx');
  DriveApp.createFile(blob);
  const url = blob.getUrl();
  const html = HtmlService.createHtmlOutput('<script>window.open("' + url + '","_blank");window.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, 'Экспорт завершен');
}

function createBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupSheet = ss.getSheetByName('Бэкап');
  if (!backupSheet) return;
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const logSheet = ss.getSheetByName('Логи');
  if (!logSheet) return;
  const logRange = logSheet.getRange(1, 1, 1, 3);
  logRange.setValues([[timestamp, 'Бэкап создан', 'Успешно']]);
  const backupRow = backupSheet.getLastRow() + 1;
  backupSheet.getRange(backupRow, 1).setValue(timestamp);
  backupSheet.getRange(backupRow, 2).setValue(ss.getUrl());
  backupSheet.getRange(backupRow, 3).setValue('Активный');
}

function createBackupTrigger() {
  ScriptApp.newTrigger('createBackup')
    .timeBased()
    .everyHours(6)
    .create();
}

function createCallLogEntry(action, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Логи');
  if (!logSheet) return;
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const user = Session.getActiveUser().getEmail();
  const row = logSheet.getLastRow() + 1;
  logSheet.getRange(row, 1).setValue(timestamp);
  logSheet.getRange(row, 2).setValue(user);
  logSheet.getRange(row, 3).setValue(action);
  logSheet.getRange(row, 4).setValue(details);
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== 'Звонки') return;
  const row = range.getRow();
  if (row === 1) return;
  const col = range.getColumn();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('Статус') + 1;
  const dateCol = headers.indexOf('Дата создания') + 1;
  const lastCallCol = headers.indexOf('Последний звонок') + 1;
  if (col === statusCol) {
    const newValue = range.getValue();
    if (newValue) {
      const now = new Date();
      sheet.getRange(row, lastCallCol).setValue(now);
      createCallLogEntry('Изменен статус', 'Строка ' + row + ': ' + newValue);
    }
  }
}

function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const statusCol = headers.indexOf('Статус');
  const statusCounts = {};
  for (let i = 1; i < values.length; i++) {
    const status = values[i][statusCol];
    if (status) {
      statusCounts[status] = (statusCounts[status] || 0) + 1;
    }
  }
  const statsSheet = ss.getSheetByName('Статистика');
  if (!statsSheet) {
    ss.insertSheet('Статистика');
  } else {
    statsSheet.clear();
  }
  const stats = statsSheet.getSheetByName('Статистика');
  stats.getRange(1, 1).setValue('Статус');
  stats.getRange(1, 2).setValue('Количество');
  let row = 2;
  for (const [status, count] of Object.entries(statusCounts)) {
    stats.getRange(row, 1).setValue(status);
    stats.getRange(row, 2).setValue(count);
    row++;
  }
  stats.getRange(1, 1, row - 1, 2).setBorder(true, true, true, true, true, true);
}

function createSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Настройки');
  if (!settingsSheet) {
    const sheet = ss.insertSheet('Настройки');
    sheet.getRange(1, 1).setValue('Параметр');
    sheet.getRange(1, 2).setValue('Значение');
    sheet.getRange(2, 1).setValue('Последний бэкап');
    sheet.getRange(2, 2).setValue(new Date().toISOString());
    sheet.getRange(3, 1).setValue('Частота бэкапа (часы)');
    sheet.getRange(3, 2).setValue(6);
    sheet.getRange(4, 1).setValue('Последний экспорт');
    sheet.getRange(4, 2).setValue('');
  }
}

function createSheetsStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = ['Звонки', 'Логи', 'Статистика', 'Бэкап', 'Настройки'];
  const existingSheets = ss.getSheets().map(s => s.getName());
  requiredSheets.forEach(name => {
    if (!existingSheets.includes(name)) {
      ss.insertSheet(name);
    }
  });
  const headers = [
    'ID', 'ФИО', 'Телефон', 'Дата создания', 'Последний звонок', 'Статус',
    'Заметки', 'Комментарий', 'Источник', 'Ответственный', 'График'
  ];
  const sheet = ss.getSheetByName('Звонки');
  if (sheet) {
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    if (lastCol < headers.length) {
      sheet.getRange(1, lastCol + 1, 1, headers.length - lastCol).setValues([headers.slice(lastCol)]);
    }
    if (lastRow < 1000) {
      sheet.getRange(lastRow + 1, 1, 1000 - lastRow, headers.length).setValues(Array(1000 - lastRow).fill(headers));
    }
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(2, 1, 1000, headers.length).setRowHeight(30);
    sheet.getRange(1, 1, 1, headers.length).setRowHeight(40);
  }
}

function createButtons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('CRM Call Tracking');
  menu.addItem('Перезвонить', 'callbackAction')
    .addItem('Добавить заметку', 'addNoteAction')
    .addItem('Сохранить', 'saveAction')
    .addToUi();
}

function callbackAction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getActiveRange();
  const row = range.getRow();
  if (row === 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const phoneCol = headers.indexOf('Телефон') + 1;
  const statusCol = headers.indexOf('Статус') + 1;
  const lastCallCol = headers.indexOf('Последний звонок') + 1;
  const phone = sheet.getRange(row, phoneCol).getValue();
  const status = sheet.getRange(row, statusCol).getValue();
  if (phone) {
    sheet.getRange(row, lastCallCol).setValue(new Date());
    sheet.getRange(row, statusCol).setValue('🟡 ПЕРЕЗВОНИТЬ');
    createCallLogEntry('Перезвонить', 'Телефон: ' + phone);
  }
}

function addNoteAction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Звонки');
  if (!sheet) return;
  const range = sheet.getActiveRange();
  const row = range.getRow();
  if (row === 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const notesCol = headers.indexOf('Заметки') + 1;
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Добавить заметку', 'Введите текст заметки:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() === ui.Button.OK) {
    const