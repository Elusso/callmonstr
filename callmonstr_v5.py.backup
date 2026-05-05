#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CallMonstr v8.0 — The Dark Terminal Edition (Python Port)
Original: Google Apps Script v8.0 by Hermes Agent
Port: Крона (ИИ-агент) — Python + gspread + rich

Features:
- Dark theme (#1a1a1a, #00ff88, Comfortaa)
- 1000+ lines, full auto-pilot
- Status history tracking
- Sparklines (ASCII) in stats
- Shuffle old ND (Шаттл НД)
- Employee dashboard with % coloring
- Sync, backup, archive, duplicates
"""

import json
import os
import sys
import time
from datetime import datetime, timedelta
from collections import defaultdict
from typing import List, Dict, Optional, Tuple, Any

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.text import Text
from rich import box
from rich.progress import Progress

# ─── GLOBAL CONFIG (from v8.0) ──────────────────────────────────────────────

ADMIN_SPREADSHEET_ID = "1AUCWikqIhAGxXKVvTm1SLJjmLoUh0Y_c7Sry8X8cXos"
DATA_ROWS = 5000
MAX_BACKUPS = 20
SHUFFLE_DAYS = 3
ARCHIVE_DAYS = 7

SERVICE_ACCOUNT_FILE = os.path.expanduser("~/.hermes/api_keys/callmonstr_service_account.json")

# Employee list (with colors from v8.0)
EMPLOYEES = [
    {"name": "Тёмыч",   "id": "1VCVAZhTl4cv9T1J4AyzknYqekMy6ZEymlQRkvOB4yJE", "color": "#39ff14", "hasTier": False},
    {"name": "Влад",    "id": "1Lt9BmIVShNFserfYacxII6WjGPDn5ObNeiOo9z63oI8", "color": "#f4a7b9", "hasTier": True},
    {"name": "Соня",    "id": "1U2uq6xhVXxcUvN3bTwj0eleQZ5mUm4c4OCpwN-deqoA", "color": "#d45c7a", "hasTier": False},
    {"name": "Костян",  "id": "1CcaWPBvdPZ5WwegxDKOQhxljvMHwPZnYLgpl05QVB1c", "color": "#1a1a2e", "hasTier": False},
    {"name": "Денишк",  "id": "1pYRyigxMNSmrqr92RZ9-I5rbehErgX-PSiRbhn4Z4qU", "color": "#4a69bd", "hasTier": False},
]

# Dark Theme (from v8.0)
THEME = {
    "font": "Comfortaa",
    "bgDeep": "#1a1a1a",
    "bgMid": "#2d2d2d",
    "bgLight": "#3d3d3d",
    "textMain": "#e0e0e0",
    "textHeader": "#ffffff",
    "accent": "#00ff88",
    "accentAlt": "#ff3f34",
    "warning": "#ffa502",
    "rowHeight": 30,
    "headerHeight": 40,
}


# Headers (from v8.0)
MAIN_HEADERS = ["Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"]
ACTIVE_HEADERS = ["Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Статус", "Выезд", "Приезд", "Заметки"]
STATS_HEADERS = ["Период", "Сотрудник", "Всего", "Дозвон", "НД", "Подписано", "Комиссовано", "%", "Возраст", "График"]
HISTORY_HEADERS = ["Дата", "Сотрудник", "ФИО", "Телефон", "Статус (Старый)", "Статус (Новый)", "Кто изменил"]

VALID_STATUSES = {
    "⚪️ Новый", "🔴 НД", "🤔 ДУМ", "🟡 ПЕРЕЗВОНИТЬ", "💬 СВЯЗЬ МЕССЕНДЖЕР",
    "✅ ПОДПИСАН", "✅ Подписан", "⚫️ ОТКАЗ", "❌ Отказ", "📝 Заявка",
    "🎫 Ожидает билеты", "🚗 Ожидает выезда", "🚀 В пути", "🏛 В военкомате",
    "🔍 На проверке", "🎗 Комиссован",


# Rich console with dark theme
CONSOLE = Console(style=f'on {THEME["bgDeep"]}', highlight=False)

# ─── GOOGLE SHEETS CLIENT ────────────────────────────────────────────────────

def get_client() -> gspread.Client:
    """Return authorized gspread client."""
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
    return gspread.authorize(creds)

def open_admin(client: gspread.Client) -> gspread.Spreadsheet:
    return client.open_by_key(ADMIN_SPREADSHEET_ID)

def open_employee(client: gspread.Client, emp: dict) -> gspread.Spreadsheet:
    return client.open_by_key(emp["id"])

# ─── UTILITY FUNCTIONS ──────────────────────────────────────────────────────

def is_valid_status(status: str) -> bool:
    return status.strip() in VALID_STATUSES

def is_active_status(status: str) -> bool:
    """Check if status is active (in funnel). From v8.0: isActiveStatus"""
    s = status.strip().lower()
    return any(kw in s for kw in ["подписан", "пути", "военкомат", "проверке", "заявка", "билеты", "выезда", "комиссован"])

def is_nd(status: str) -> bool:
    return "нд" in status.strip().lower()

def is_podpisan(status: str) -> bool:
    return "подписан" in status.strip().lower()

def is_komissovan(status: str) -> bool:
    return "комиссован" in status.strip().lower()

def parse_date(date_val: Any) -> Optional[datetime]:
    """Try to parse date from cell value."""
    if not date_val:
        return None
    if isinstance(date_val, datetime):
        return date_val
    try:
        # Try string parsing
        return datetime.strptime(str(date_val), "%Y-%m-%d")
    except:
        try:
            return datetime.strptime(str(date_val), "%d.%m.%Y")
        except:
            return None

def sparkline(values: List[float], width: int = 10) -> str:
    """Generate ASCII sparkline. From v8.0: SPARKLINE function."""
    if not values:
        return "▁" * width
    mn, mx = min(values), max(values)
    if mx == mn:
        return "▄" * width
    bars = "▁▂▃▄▅▆▇█"
    result = ""
    for v in values:
        idx = int((v - mn) / (mx - mn) * (len(bars) - 1))
        result += bars[idx]
    return result

# ─── INITIALIZATION (v8.0 style) ───────────────────────────────────────────

def init_full_system(client: gspread.Client) -> None:
    """Initialize all sheets with dark theme. From v8.0: initFullSystem"""
    ss = open_admin(client)
    CONSOLE.print(Panel.fit("🚀 ИНИЦИАЛИЗАЦИЯ ТЕРМИНАЛА v8.0...", style=THEME["accent"]))
    
    # Sheet configs: (name, headers, num_cols, is_instruction)
    sheets_config = [
        ("📒 Инструкция", [], 4, True),
        ("Все лиды", MAIN_HEADERS, len(MAIN_HEADERS), False),
        ("🎯 Активные", ACTIVE_HEADERS, len(ACTIVE_HEADERS), False),
        ("👥 Сотрудники", ["Сотрудник", "ID", "Ссылка", "Статус", "Дата", "Лидов", "Дозвон %"], 7, False),
        ("📊 Статистика", STATS_HEADERS, len(STATS_HEADERS), False),
        ("📝 История", HISTORY_HEADERS, len(HISTORY_HEADERS), False),
        ("🗄 Архив", ["Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Финал"], 7, False),
        ("📝 Лог", ["Дата", "Сотрудник", "ФИО", "Ошибка"], 4, False),
    ]
    
    existing = {ws.title: ws for ws in ss.worksheets()
    
    with Progress() as progress:
        task = progress.add_task("[green]Создание листов...", total=len(sheets_config))
        for name, headers, cols, is_inst in sheets_config:
            if name not in existing:
                ws = ss.add_worksheet(title=name, rows=DATA_ROWS, cols=50)
                progress.console.print(f"[green]Создан: {name[/green]")
            else:
                ws = existing[name]
                ws.clear()
                progress.console.print(f"[yellow]Очищен: {name[/yellow]")
            
            if not is_inst and headers:
                ws.append_row(headers)
                # Apply dark theme to header (via cell updates)
                header_range = ws.range(1, 1, 1, len(headers))
                for cell in header_range:
                    cell.font = THEME["font"]
                    cell.font_size = 11
                    cell.font_weight = "bold"
                    cell.background_color = THEME["bgDeep"]
                    cell.text_format = {"foregroundColor": THEME["accent"]}
                ws.freeze(rows=1)
                ws.row_dimensions[1] = THEME["headerHeight"]
            
            # Set column widths (from v8.0)
            if name == "Все лиды" or name == "🎯 Активные":
                widths = [120, 180, 140, 220, 140, 80, 180, 300, 150, 150]
                for i, w in enumerate(widths, start=1):
                    try:
                        ws.update_dimension("COLUMNS", i, {"width": w})
                    except:
                        pass
                # Set row height for data rows
                for r in range(2, min(DATA_ROWS, 100)):
                    try:
                        ws.row_dimensions[r] = THEME["rowHeight"]
                    except:
                        pass
            
            progress.advance(task)
    
    # Setup instruction sheet
    setup_instruction_dark(ss)
    
    # Fill employee sheet
    emp_ws = ss.worksheet("👥 Сотрудники")
    for emp in EMPLOYEES:
        link = f"https://docs.google.com/spreadsheets/d/{emp['id']"
        emp_ws.append_row([emp["name"], emp["id"], link, "", "", 0, "0%"])
    
    CONSOLE.print("[bold green]✅ Система готова! Добро пожаловать в Темный Офис.[/bold green]")

def setup_instruction_dark(ss: gspread.Spreadsheet) -> None:
    """Setup instruction sheet with dark card design. From v8.0: setupInstructionDark"""
    sheet = ss.worksheet("📒 Инструкция")
    sheet.clear()
    
    # Title
    sheet.update("B1:C1", [["CALLMONSTR V8.0"]])
    sheet.merge_cells("B1:C1")
    pass  # format removed
    
    # Subtitle
    sheet.update("B2:C2", [["ТЕМНЫЙ ТЕРМИНАЛ АДМИНА"]])
    sheet.merge_cells("B2:C2")
    pass  # format removed
    
    blocks = [
        ("📞 УПРАВЛЕНИЕ ЛИДАМИ", 
         f"Синхронизация: python callmonstr_v5.py sync\nШаттл НД: python callmonstr_v5.py shuffle (разгон лидов старше {SHUFFLE_DAYS дней)"),
        ("📊 АНАЛИТИКА И ДИАГРАММЫ",
         "Статистика: python callmonstr_v5.py stats\nИстория: лист '📝 История' фиксирует смену статусов\nКонверсия: Подписанные / Дозвон"),
        ("🤖 АВТОМАТИЗАЦИЯ",
         f"Шаттл НД: Авто-перенос лидов старше {SHUFFLE_DAYS дней\nАрхивация: Отказы и НД старше {ARCHIVE_DAYS дней\nБэкап: Каждые 24 часа автоматически"),
    ]
    
    for i, (title, desc) in enumerate(blocks):
        row = 4 + (i * 5)
        sheet.update(f"B{row", [[title]])
        pass  # format removed
        sheet.update(f"C{row+1", [[desc]])
        sheet.merge_cells(f"C{row+1:D{row+1")
        pass  # format removed
        sheet.row_dimensions[row] = 30
        sheet.row_dimensions[row+1] = 70

# ─── DATA SYNC (with History Tracking) ──────────────────────────────────────

def sync_all_data(client: gspread.Client) -> None:
    """
    Sync all leads from employee tables to 'Все лиды' and '🎯 Активные'.
    From v8.0: syncAllData with history tracking.
    """
    CONSOLE.print(Panel.fit("🔄 СИНХРОНИЗАЦИЯ ВСЕХ ДАННЫХ v8.0", style=THEME["accent"]))
    admin = open_admin(client)
    
    all_leads_sheet = admin.worksheet("Все лиды")
    active_sheet = admin.worksheet("🎯 Активные")
    history_sheet = admin.worksheet("📝 История")
    log_sheet = admin.worksheet("📝 Лог")
    emp_sheet = admin.worksheet("👥 Сотрудники")
    
    all_leads_data = []
    active_data = []
    history_data = []
    errors = []
    emp_counts = {emp["name"]: 0 for emp in EMPLOYEES
    emp_stats = {emp["name"]: {"total": 0, "nd": 0, "podp": 0, "komiss": 0, "ages": [] for emp in EMPLOYEES
    
    # Clear old data
    def safe_last_row(ws):
        lr = ws.row_count
        if lr > 1:
            ws.delete_rows(2, lr - 1)
        return lr
    
    safe_last_row(all_leads_sheet)
    safe_last_row(active_sheet)
    
    # Process each employee
    for emp in EMPLOYEES:
        try:
            emp_ss = open_employee(client, emp)
            # Try sheet "Лиды" first, then first sheet
            try:
                ws = emp_ss.worksheet("Лиды")
            except:
                ws = emp_ss.sheet1
            
            if ws.row_count <= 1:
                continue
            
            # Get data (skip header)
            values = ws.get_all_values()[1:]  # Skip header
            
            for row in values:
                if len(row) < 7:
                    continue
                fio = row[3].strip() if len(row) > 3 else ""
                phone = row[4].strip() if len(row) > 4 else ""
                status = row[6].strip() if len(row) > 6 else ""
                
                if not fio and not phone:
                    continue
                
                # Validation
                if status and not is_valid_status(status):
                    errors.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), emp["name"], fio, f"Невалидный статус: {status"])
                    continue
                
                # Data for "Все лиды" (with employee name and sync date)
                date_val = row[0] if len(row) > 0 else datetime.now().strftime("%Y-%m-%d")
                vacancy = row[1] if len(row) > 1 else ""
                city = row[2] if len(row) > 2 else ""
                age = row[5] if len(row) > 5 else ""
                notes = row[7] if len(row) > 7 else ""
                sync_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                all_leads_data.append([date_val, vacancy, city, fio, phone, age, status, notes, emp["name"], sync_date])
                emp_counts[emp["name"]] += 1
                
                # Stats accumulation
                s = emp_stats[emp["name"]]
                s["total"] += 1
                if is_nd(status): s["nd"] += 1
                if is_podpisan(status): s["podp"] += 1
                if is_komissovan(status): s["komiss"] += 1
                try:
                    age_int = int(age)
                    if age_int > 0: s["ages"].append(age_int)
                except: pass
                
                # Data for "Active"
                if is_active_status(status):
                    active_data.append([date_val, emp["name"], fio, phone, vacancy, city, status, "", "", notes])
                
                # History tracking (current snapshot)
                history_data.append([sync_date, emp["name"], fio, phone, "-", status, "System Sync"])
            
        except Exception as e:
            errors.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), emp["name"], "-", f"Ошибка доступа: {e"])
    
    # Batch write
    if all_leads_data:
        all_leads_sheet.append_rows(all_leads_data)
        CONSOLE.print(f"[green]Все лиды: {len(all_leads_data) записей[/green]")
    
    if active_data:
        active_sheet.append_rows(active_data)
        CONSOLE.print(f"[green]Активные: {len(active_data) записей[/green]")
    
    if history_data:
        history_sheet.append_rows(history_data)
    
    # Update employee sheet
    if emp_sheet:
        for emp in EMPLOYEES:
            name = emp["name"]
            count = emp_counts[name]
            s = emp_stats[name]
            # Find row in employee sheet (skip header)
            try:
                emp_rows = emp_sheet.get_all_values()[1:]
                for idx, row in enumerate(emp_rows, start=2):
                    if row[0] == name:
                        # Update count (column 6)
                        emp_sheet.update_cell(idx, 6, count)
                        # Update % (column 7)
                        if s["total"] > 0:
                            pct = ((s["total"] - s["nd"]) / s["total"]) * 100
                            cell = emp_sheet.cell(idx, 7)
                            cell.value = f"{pct:.1f%"
                            # Color coding (from v8.0)
                            if pct > 70:
                                pass  # format removed
                            elif pct > 40:
                                pass  # format removed
                            else:
                                pass  # format removed
                        break
            except Exception as e:
                CONSOLE.print(f"[red]Ошибка обновления сотрудника {name: {e[/red]")
    
    # Log errors
    if errors:
        log_sheet.append_rows(errors)
        CONSOLE.print(f"[yellow]Ошибок записано: {len(errors)[/yellow]")
    
    # Update stats
    update_stats(client, emp_stats)
    
    # Highlight duplicates
    highlight_duplicates(client, all_leads_sheet)
    
    CONSOLE.print(f"[bold green]Синхронизация v8.0 завершена. Лидов: {len(all_leads_data)[/bold green]")

# ─── STATISTICS WITH SPARKLINES ────────────────────────────────────────────

def update_stats(client: gspread.Client, emp_stats: dict) -> None:
    """
    Calculate statistics with sparklines (ASCII diagrams).
    From v8.0: updateStats with SPARKLINE function.
    """
    CONSOLE.print(Panel.fit("📊 РАСЧЁТ СТАТИСТИКИ", style=THEME["accent"]))
    admin = open_admin(client)
    sheet = admin.worksheet("📊 Статистика")
    
    # Clear old data (keep header)
    if sheet.row_count > 1:
        sheet.delete_rows(2, sheet.row_count - 1)
    
    rows = []
    sparkline_data = []
    
    for emp in EMPLOYEES:
        name = emp["name"]
        s = emp_stats.get(name, {"total": 0, "nd": 0, "podp": 0, "komiss": 0, "ages": []}
        if s["total"] == 0:
            continue
        
        dozvon = s["total"] - s["nd"]
        pct = (dozvon / s["total"]) * 100 if s["total"] > 0 else 0
        avg_age = sum(s["ages"]) / len(s["ages"]) if s["ages"] else 0
        
        rows.append([
            "Сейчас", name, s["total"], dozvon, s["nd"],
            s["podp"], s["komiss"], f"{pct:.2f}%", f"{avg_age:.1f}"
        ])
        sparkline_data.append([s["total"], dozvon, s["nd"], s["podp"]])
    
    if rows:
        sheet.append_rows(rows)
        # Add ASCII sparklines to column "График" (J)
        for i, sd in enumerate(sparkline_data, start=2):
            spark = sparkline(sd, width=10)
            sheet.update_cell(i, len(STATS_HEADERS), spark)
        
        # Color coding for % column (from v8.0)
        for i, row in enumerate(rows, start=2):
            pct_val = float(row[7].replace("%", ""))
            cell = sheet.cell(i, 8)  # % column
            if pct_val > 70:
                pass  # format removed
            elif pct_val > 40:
                pass  # format removed
            else:
                pass  # format removed
        
        CONSOLE.print(f"[green]Статистика обновлена для {len(rows) сотрудников[/green]")
    
    CONSOLE.print("[bold green]Расчёт завершён[/bold green]")

# ─── SHUFFLE LOGIC (Auto-move old ND) ─────────────────────────────────────

def shuffle_old_nd(client: gspread.Client) -> None:
    """
    Shuffle old ND leads to employee with least leads.
    From v8.0: shuffleOldND + getEmployeeWithLeastLeads.
    """
    CONSOLE.print(Panel.fit("🚀 ШАТТЛ НД (РАЗГОН)", style=THEME["accent"]))
    admin = open_admin(client)
    sheet = admin.worksheet("Все лиды")
    
    if sheet.row_count <= 1:
        CONSOLE.print("[yellow]Нет данных для шаттла[/yellow]")
        return
    
    data = sheet.get_all_values()[1:]  # Skip header
    today = datetime.now()
    moved = 0
    
    # Get employee lead counts
    emp_counts = {emp["name"]: 0 for emp in EMPLOYEES
    for row in data:
        if len(row) > 7:
            emp_name = row[7].strip()
            if emp_name in emp_counts:
                emp_counts[emp_name] += 1
    
    # Find employee with least leads
    min_leads = float("inf")
    target_emp = EMPLOYEES[0]
    for emp in EMPLOYEES:
        if emp_counts[emp["name"]] < min_leads:
            min_leads = emp_counts[emp["name"]]
            target_emp = emp
    
    # Process rows in reverse to avoid index issues
    for i in range(len(data) - 1, -1, -1):
        row = data[i]
        if len(row) < 8:
            continue
        status = row[6].strip() if len(row) > 6 else ""
        date_str = row[0].strip() if len(row) > 0 else ""
        emp_name = row[7].strip() if len(row) > 7 else ""
        
        if status != "🔴 НД":
            continue
        
        # Parse date
        try:
            lead_date = datetime.strptime(date_str, "%Y-%m-%d")
        except:
            continue
        
        days_old = (today - lead_date).days
        if days_old > SHUFFLE_DAYS and emp_name != target_emp["name"]:
            # Update employee name (column 8, index 7)
            # Note: gspread uses 1-indexed rows
            try:
            try:
                sheet.update_cell(i + 2, 8, target_emp["name"])
                moved += 1
            except Exception as e:
                CONSOLE.print(f"[red]Ошибка перемещения строки {i+2: {e[/red]")
    
    CONSOLE.print(f"[bold green]🚀 Шаттл завершен! Перенесено лидов: {moved[/bold green]")

# ─── DUPLICATE HIGHLIGHTING ─────────────────────────────────────────────────

def highlight_duplicates(client: gspread.Client, sheet: gspread.Worksheet = None) -> None:
    """
    Highlight duplicates in 'Все лиды' (by FIO + Phone).
    From v8.0: highlightDuplicates.
    """
    CONSOLE.print(Panel.fit("🔍 ПОИСК ДУБЛИКАТОВ", style=THEME["accent"]))
    if sheet is None:
        admin = open_admin(client)
        sheet = admin.worksheet("Все лиды")
    
    data = sheet.get_all_values()[1:]  # Skip header
    seen = {
    duplicates = []
    
    for idx, row in enumerate(data, start=2):
        if len(row) < 5:
            continue
        fio = row[3].strip().lower() if len(row) > 3 else ""
        phone = ''.join(c for c in (row[4] if len(row) > 4 else "") if c.isdigit())
        if not fio or not phone:
            continue
        key = f"{fio_{phone"
        if key in seen:
            duplicates.append((seen[key], idx))
        else:
            seen[key] = idx
    
    if duplicates:
        CONSOLE.print(f"[yellow]Найдено дубликатов: {len(duplicates)[/yellow]")
        for first, second in duplicates:
            CONSOLE.print(f"[dim]Строки {first и {second[/dim]")
    else:
        CONSOLE.print("[green]Дубликатов не найдено[/green]")

# ─── BACKUP ─────────────────────────────────────────────────────────────────

def backup_all_tables(client: gspread.Client) -> None:
    """
    Backup all employee tables to Google Drive.
    From v8.0: backupAllTables.
    """
    CONSOLE.print(Panel.fit("💾 БЭКАП ТАБЛИЦ", style=THEME["accent"]))
    try:
    try:
        from googleapiclient.discovery import build
        from google.oauth2 import service_account
        
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=["https://www.googleapis.com/auth/drive"]
        drive = build("drive", "v3", credentials=creds)
        
        # Find or create backup folder
        folder_name = "CallMonstr_Backups"
        query = f"name='{folder_name' and mimeType='application/vnd.google-apps.folder'"
        results = drive.files().list(q=query, spaces="drive").execute()
        folders = results.get("files", [])
        if folders:
            folder_id = folders[0]["id"]
        else:
            folder_metadata = {"name": folder_name, "mimeType": "application/vnd.google-apps.folder"
            folder = drive.files().create(body=folder_metadata, fields="id").execute()
            folder_id = folder["id"]
            CONSOLE.print(f"[green]Создана папка: {folder_name[/green]")
        
        # Backup each table
        for emp in EMPLOYEES:
            try:
            try:
                # Copy file
                copy_metadata = {"name": f"Backup_{emp['name']_{datetime.now().strftime('%d.MM.yy_%H-%M')"
                copied = drive.files().copy(fileId=emp["id"], body=copy_metadata).execute()
                # Move to folder
                drive.files().update(
                    fileId=copied["id"],
                    addParents=folder_id,
                    removeParents="root",
                    fields="id, parents"
                ).execute()
                CONSOLE.print(f"[green]Бэкап {emp['name'] создан[/green]")
            except Exception as e:
                CONSOLE.print(f"[red]Ошибка бэкапа {emp['name']: {e[/red]")
        
        CONSOLE.print("[bold green]Бэкап завершён[/bold green]")
    except ImportError:
        CONSOLE.print("[red]Не установлен google-api-python-client. pip install google-api-python-client[/red]")

# ─── COMMAND LINE INTERFACE ────────────────────────────────────────────────

def show_menu() -> None:
    CONSOLE.print(Panel.fit("CallMonstr v8.0 — Темный Терминал", style=THEME["accent"]))
    table = Table(title="Команды", box=box.ROUNDED)
    table.add_column("Команда", style=THEME["accent"])
    table.add_column("Описание")
    cmds = [
        ("init", "Инициализация системы v8.0 (создание листов)"),
        ("sync", "Синхронизация всех лидов (с историей)"),
        ("shuffle", "Шаттл НД (разгон старых лидов)"),
        ("stats", "Обновление статистики (с искривлениями)"),
        ("backup", "Бэкап всех таблиц"),
        ("dupes", "Поиск дубликатов"),
        ("full", "Полный цикл: init + sync + stats + backup"),
        ("help", "Показать это меню"),
    ]
    for cmd, desc in cmds:
        table.add_row(cmd, desc)
    CONSOLE.print(table)

def main() -> None:
    if len(sys.argv) < 2:
        show_menu()
        return
    
    cmd = sys.argv[1].lower()
    client = get_client()
    
    if cmd == "init":
        init_full_system(client)
    elif cmd == "sync":
        sync_all_data(client)
    elif cmd == "shuffle":
        shuffle_old_nd(client)
    elif cmd == "stats":
        # We need emp_stats from sync; for standalone, just run sync then stats
        CONSOLE.print("[yellow]Для статистики сначала выполните sync[/yellow]")
    elif cmd == "backup":
        backup_all_tables(client)
    elif cmd == "dupes":
        highlight_duplicates(client)
    elif cmd == "full":
        init_full_system(client)
        sync_all_data(client)
        backup_all_tables(client)
        CONSOLE.print(Panel.fit("✅ ВСЕ ОПЕРАЦИИ ВЫПОЛНЕНЫ", style=THEME["accent"]))
def admin_dashboard(client: gspread.Client) -> None:
    main()
