     1|#!/usr/bin/env python3
     2|# -*- coding: utf-8 -*-
     3|"""
     4|CallMonstr v8.0 — The Dark Terminal Edition (Python Port)
     5|Original: Google Apps Script v8.0 by Hermes Agent
     6|Port: Крона (ИИ-агент) — Python + gspread + rich
     7|
     8|Features:
     9|- Dark theme (#1a1a1a, #00ff88, Comfortaa)
    10|- 1000+ lines, full auto-pilot
    11|- Status history tracking
    12|- Sparklines (ASCII) in stats
    13|- Shuffle old ND (Шаттл НД)
    14|- Employee dashboard with % coloring
    15|- Sync, backup, archive, duplicates
    16|"""
    17|
    18|import json
    19|import os
    20|import sys
    21|import time
    22|from datetime import datetime, timedelta
    23|from collections import defaultdict
    24|from typing import List, Dict, Optional, Tuple, Any
    25|
    26|import gspread
    27|from oauth2client.service_account import ServiceAccountCredentials
    28|from rich.console import Console
    29|from rich.table import Table
    30|from rich.panel import Panel
    31|from rich.text import Text
    32|from rich import box
    33|from rich.progress import Progress
    34|
    35|# ─── GLOBAL CONFIG (from v8.0) ──────────────────────────────────────────────
    36|
    37|ADMIN_SPREADSHEET_ID = "1AUCWikqIhAGxXKVvTm1SLJjmLoUh0Y_c7Sry8X8cXos"
    38|DATA_ROWS = 5000
    39|MAX_BACKUPS = 20
    40|SHUFFLE_DAYS = 3
    41|ARCHIVE_DAYS = 7
    42|
    43|SERVICE_ACCOUNT_FILE = os.path.expanduser("~/.hermes/api_keys/callmonstr_service_account.json")
    44|
    45|# Employee list (with colors from v8.0)
    46|EMPLOYEES = [
    47|    {"name": "Тёмыч",   "id": "1VCVAZhTl4cv9T1J4AyzknYqekMy6ZEymlQRkvOB4yJE", "color": "#39ff14", "hasTier": False},
    48|    {"name": "Влад",    "id": "1Lt9BmIVShNFserfYacxII6WjGPDn5ObNeiOo9z63oI8", "color": "#f4a7b9", "hasTier": True},
    49|    {"name": "Соня",    "id": "1U2uq6xhVXxcUvN3bTwj0eleQZ5mUm4c4OCpwN-deqoA", "color": "#d45c7a", "hasTier": False},
    50|    {"name": "Костян",  "id": "1CcaWPBvdPZ5WwegxDKOQhxljvMHwPZnYLgpl05QVB1c", "color": "#1a1a2e", "hasTier": False},
    51|    {"name": "Денишк",  "id": "1pYRyigxMNSmrqr92RZ9-I5rbehErgX-PSiRbhn4Z4qU", "color": "#4a69bd", "hasTier": False},
    52|]
    53|
    54|# Dark Theme (from v8.0)
    55|THEME = {
    56|    "font": "Comfortaa",
    57|    "bgDeep": "#1a1a1a",
    58|    "bgMid": "#2d2d2d",
    59|    "bgLight": "#3d3d3d",
    60|    "textMain": "#e0e0e0",
    61|    "textHeader": "#ffffff",
    62|    "accent": "#00ff88",
    63|    "accentAlt": "#ff3f34",
    64|    "warning": "#ffa502",
    65|    "rowHeight": 30,
    66|    "headerHeight": 40,
    67|}
    68|
    69|# Headers (from v8.0)
    70|MAIN_HEADERS = ["Дата", "Вакансия", "Город", "ФИО", "Телефон", "Возраст", "Статус", "Заметки"]
    71|ACTIVE_HEADERS = ["Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Статус", "Выезд", "Приезд", "Заметки"]
    72|STATS_HEADERS = ["Период", "Сотрудник", "Всего", "Дозвон", "НД", "Подписано", "Комиссовано", "%", "Возраст", "График"]
    73|HISTORY_HEADERS = ["Дата", "Сотрудник", "ФИО", "Телефон", "Статус (Старый)", "Статус (Новый)", "Кто изменил"]
    74|
    75|VALID_STATUSES = {
    76|    "[NEW] Новый", "[ND] НД", "[DUM] ДУМ", "[CALLBACK] ПЕРЕЗВОНИТЬ", "[MSG] СВЯЗЬ МЕССЕНДЖЕР",
    77|    "[OK] ПОДПИСАН", "[OK] Подписан", "[DECLINED] ОТКАЗ", "[FAIL] Отказ", "[APP] Заявка",
    78|    "[TICKET] Ожидает билеты", "[DEPART] Ожидает выезда", "[ROAD] В пути", "[MIL] В военкомате",
    79|    "[CHECK] На проверке", "[KOM] Комиссован",
    80|}
    81|
    82|# Rich console with dark theme
    83|CONSOLE = Console(style=f"on {THEME['bgDeep']}", highlight=False)
    84|
    85|# ─── GOOGLE SHEETS CLIENT ────────────────────────────────────────────────────
    86|
    87|def get_client() -> gspread.Client:
    88|    """Return authorized gspread client."""
    89|    scope = [
    90|        "https://spreadsheets.google.com/feeds",
    91|        "https://www.googleapis.com/auth/spreadsheets",
    92|        "https://www.googleapis.com/auth/drive",
    93|    ]
    94|    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
    95|    return gspread.authorize(creds)
    96|
    97|def open_admin(client: gspread.Client) -> gspread.Spreadsheet:
    98|    return client.open_by_key(ADMIN_SPREADSHEET_ID)
    99|
   100|def open_employee(client: gspread.Client, emp: dict) -> gspread.Spreadsheet:
   101|    return client.open_by_key(emp["id"])
   102|
   103|# ─── UTILITY FUNCTIONS ──────────────────────────────────────────────────────
   104|
   105|def is_valid_status(status: str) -> bool:
   106|    return status.strip() in VALID_STATUSES
   107|
   108|def is_active_status(status: str) -> bool:
   109|    """Check if status is active (in funnel). From v8.0: isActiveStatus"""
   110|    s = status.strip().lower()
   111|    return any(kw in s for kw in ["подписан", "пути", "военкомат", "проверке", "заявка", "билеты", "выезда", "комиссован"])
   112|
   113|def is_nd(status: str) -> bool:
   114|    return "нд" in status.strip().lower()
   115|
   116|def is_podpisan(status: str) -> bool:
   117|    return "подписан" in status.strip().lower()
   118|
   119|def is_komissovan(status: str) -> bool:
   120|    return "комиссован" in status.strip().lower()
   121|
   122|def parse_date(date_val: Any) -> Optional[datetime]:
   123|    """Try to parse date from cell value."""
   124|    if not date_val:
   125|        return None
   126|    if isinstance(date_val, datetime):
   127|        return date_val
   128|    try:
   129|        # Try string parsing
   130|        return datetime.strptime(str(date_val), "%Y-%m-%d")
   131|    except:
   132|        try:
   133|            return datetime.strptime(str(date_val), "%d.%m.%Y")
   134|        except:
   135|            return None
   136|
   137|def sparkline(values: List[float], width: int = 10) -> str:
   138|    """Generate ASCII sparkline. From v8.0: SPARKLINE function."""
   139|    if not values:
   140|        return "_" * width
   141|    mn, mx = min(values), max(values)
   142|    if mx == mn:
   143|        return "_" * width
   144|    bars = "_______#"
   145|    result = ""
   146|    for v in values:
   147|        idx = int((v - mn) / (mx - mn) * (len(bars) - 1))
   148|        result += bars[idx]
   149|    return result
   150|
   151|# ─── INITIALIZATION (v8.0 style) ───────────────────────────────────────────
   152|
   153|def init_full_system(client: gspread.Client) -> None:
   154|    """Initialize all sheets with dark theme. From v8.0: initFullSystem"""
   155|    ss = open_admin(client)
   156|    CONSOLE.print(Panel.fit("[INIT] ИНИЦИАЛИЗАЦИЯ ТЕРМИНАЛА v8.0...", style=THEME["accent"]))
   157|    
   158|    # Sheet configs: (name, headers, num_cols, is_instruction)
   159|    sheets_config = [
   160|        ("[INSTR] Инструкция", [], 4, True),
   161|        ("Все лиды", MAIN_HEADERS, len(MAIN_HEADERS), False),
   162|        ("[ACTIVE] Активные", ACTIVE_HEADERS, len(ACTIVE_HEADERS), False),
   163|        ("[EMP] Сотрудники", ["Сотрудник", "ID", "Ссылка", "Статус", "Дата", "Лидов", "Дозвон %"], 7, False),
   164|        ("[STATS] Статистика", STATS_HEADERS, len(STATS_HEADERS), False),
   165|        ("[APP] История", HISTORY_HEADERS, len(HISTORY_HEADERS), False),
   166|        ("[ARCH] Архив", ["Дата", "Сотрудник", "ФИО", "Телефон", "Вакансия", "Город", "Финал"], 7, False),
   167|        ("[APP] Лог", ["Дата", "Сотрудник", "ФИО", "Ошибка"], 4, False),
   168|    ]
   169|    
   170|    existing = {ws.title: ws for ws in ss.worksheets()}
   171|    
   172|    with Progress() as progress:
   173|        task = progress.add_task("[green]Создание листов...", total=len(sheets_config))
   174|        for name, headers, cols, is_inst in sheets_config:
   175|            if name not in existing:
   176|                ws = ss.add_worksheet(title=name, rows=DATA_ROWS, cols=50)
   177|                progress.console.print(f"[green]Создан: {name}[/green]")
   178|            else:
   179|                ws = existing[name]
   180|                ws.clear()
   181|                progress.console.print(f"[yellow]Очищен: {name}[/yellow]")
   182|            
   183|            if not is_inst and headers:
   184|                ws.append_row(headers)
   185|                # Apply dark theme to header (via cell updates)
   186|                header_range = ws.range(1, 1, 1, len(headers))
   187|                for cell in header_range:
   188|                    cell.font = THEME["font"]
   189|                    cell.font_size = 11
   190|                    cell.font_weight = "bold"
   191|                    cell.background_color = THEME["bgDeep"]
   192|                    cell.text_format = {"foregroundColor": THEME["accent"]}
   193|                ws.freeze(rows=1)
   194|                ws.row_dimensions[1] = THEME["headerHeight"]
   195|            
   196|            # Set column widths (from v8.0)
   197|            if name == "Все лиды" or name == "[ACTIVE] Активные":
   198|                widths = [120, 180, 140, 220, 140, 80, 180, 300, 150, 150]
   199|                for i, w in enumerate(widths, start=1):
   200|                    try:
   201|                        ws.update_dimension("COLUMNS", i, {"width": w})
   202|                    except:
   203|                        pass
   204|                # Set row height for data rows
   205|                for r in range(2, min(DATA_ROWS, 100)):
   206|                    try:
   207|                        ws.row_dimensions[r] = THEME["rowHeight"]
   208|                    except:
   209|                        pass
   210|            
   211|            progress.advance(task)
   212|    
   213|    # Setup instruction sheet
   214|    setup_instruction_dark(ss)
   215|    
   216|    # Fill employee sheet
   217|    emp_ws = ss.worksheet("[EMP] Сотрудники")
   218|    for emp in EMPLOYEES:
   219|        link = f"https://docs.google.com/spreadsheets/d/{emp['id']}"
   220|        emp_ws.append_row([emp["name"], emp["id"], link, "", "", 0, "0%"])
   221|    
   222|    CONSOLE.print("[bold green][OK] Система готова! Добро пожаловать в Темный Офис.[/bold green]")
   223|
   224|def setup_instruction_dark(ss: gspread.Spreadsheet) -> None:
   225|    """Setup instruction sheet with dark card design. From v8.0: setupInstructionDark"""
   226|    sheet = ss.worksheet("[INSTR] Инструкция")
   227|    sheet.clear()
   228|    
   229|    # Title
   230|    sheet.update("B1:C1", [["CALLMONSTR V8.0"]])
   231|    sheet.merge_cells("B1:C1")
   232|    sheet.format("B1:C1", {
   233|        "fontFamily": THEME["font"], "fontSize": 20, "fontWeight": "bold",
   234|        "backgroundColor": THEME["bgDeep"], "textFormat": {"foregroundColor": THEME["accent"]},
   235|        "horizontalAlignment": "center",
   236|    })
   237|    
   238|    # Subtitle
   239|    sheet.update("B2:C2", [["ТЕМНЫЙ ТЕРМИНАЛ АДМИНА"]])
   240|    sheet.merge_cells("B2:C2")
   241|    sheet.format("B2:C2", {
   242|        "fontFamily": THEME["font"], "fontSize": 12,
   243|        "backgroundColor": THEME["bgMid"], "textFormat": {"foregroundColor": THEME["textMain"]},
   244|        "horizontalAlignment": "center",
   245|    })
   246|    
   247|    blocks = [
   248|        ("[PHONE] УПРАВЛЕНИЕ ЛИДАМИ", 
   249|         f"Синхронизация: python callmonstr_v5.py sync\nШаттл НД: python callmonstr_v5.py shuffle (разгон лидов старше {SHUFFLE_DAYS} дней)"),
   250|        ("[STATS] АНАЛИТИКА И ДИАГРАММЫ",
   251|         "Статистика: python callmonstr_v5.py stats\nИстория: лист '[APP] История' фиксирует смену статусов\nКонверсия: Подписанные / Дозвон"),
   252|        ("[AUTO] АВТОМАТИЗАЦИЯ",
   253|         f"Шаттл НД: Авто-перенос лидов старше {SHUFFLE_DAYS} дней\nАрхивация: Отказы и НД старше {ARCHIVE_DAYS} дней\nБэкап: Каждые 24 часа автоматически"),
   254|    ]
   255|    
   256|    for i, (title, desc) in enumerate(blocks):
   257|        row = 4 + (i * 5)
   258|        sheet.update(f"B{row}", [[title]])
   259|        sheet.format(f"B{row}", {
   260|            "fontFamily": THEME["font"], "fontSize": 13, "fontWeight": "bold",
   261|            "backgroundColor": THEME["bgLight"], "textFormat": {"foregroundColor": THEME["accent"]},
   262|        })
   263|        sheet.update(f"C{row+1}", [[desc]])
   264|        sheet.merge_cells(f"C{row+1}:D{row+1}")
   265|        sheet.format(f"C{row+1}", {
   266|            "fontFamily": THEME["font"], "fontSize": 10,
   267|            "backgroundColor": THEME["bgMid"], "textFormat": {"foregroundColor": THEME["textMain"]},
   268|            "wrapText": True,
   269|        })
   270|        sheet.row_dimensions[row] = 30
   271|        sheet.row_dimensions[row+1] = 70
   272|
   273|# ─── DATA SYNC (with History Tracking) ──────────────────────────────────────
   274|
   275|def sync_all_data(client: gspread.Client) -> None:
   276|    """
   277|    Sync all leads from employee tables to 'Все лиды' and '[ACTIVE] Активные'.
   278|    From v8.0: syncAllData with history tracking.
   279|    """
   280|    CONSOLE.print(Panel.fit("[SYNC] СИНХРОНИЗАЦИЯ ВСЕХ ДАННЫХ v8.0", style=THEME["accent"]))
   281|    admin = open_admin(client)
   282|    
   283|    all_leads_sheet = admin.worksheet("Все лиды")
   284|    active_sheet = admin.worksheet("[ACTIVE] Активные")
   285|    history_sheet = admin.worksheet("[APP] История")
   286|    log_sheet = admin.worksheet("[APP] Лог")
   287|    emp_sheet = admin.worksheet("[EMP] Сотрудники")
   288|    
   289|    all_leads_data = []
   290|    active_data = []
   291|    history_data = []
   292|    errors = []
   293|    emp_counts = {emp["name"]: 0 for emp in EMPLOYEES}
   294|    emp_stats = {emp["name"]: {"total": 0, "nd": 0, "podp": 0, "komiss": 0, "ages": []} for emp in EMPLOYEES}
   295|    
   296|    # Clear old data
   297|    def safe_last_row(ws):
   298|        lr = ws.row_count
   299|        if lr > 1:
   300|            ws.delete_rows(2, lr - 1)
   301|        return lr
   302|    
   303|    safe_last_row(all_leads_sheet)
   304|    safe_last_row(active_sheet)
   305|    
   306|    # Process each employee
   307|    for emp in EMPLOYEES:
   308|        try:
   309|            emp_ss = open_employee(client, emp)
   310|            # Try sheet "Лиды" first, then first sheet
   311|            try:
   312|                ws = emp_ss.worksheet("Лиды")
   313|            except:
   314|                ws = emp_ss.sheet1
   315|            
   316|            if ws.row_count <= 1:
   317|                continue
   318|            
   319|            # Get data (skip header)
   320|            values = ws.get_all_values()[1:]  # Skip header
   321|            
   322|            for row in values:
   323|                if len(row) < 7:
   324|                    continue
   325|                fio = row[3].strip() if len(row) > 3 else ""
   326|                phone = row[4].strip() if len(row) > 4 else ""
   327|                status = row[6].strip() if len(row) > 6 else ""
   328|                
   329|                if not fio and not phone:
   330|                    continue
   331|                
   332|                # Validation
   333|                if status and not is_valid_status(status):
   334|                    errors.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), emp["name"], fio, f"Невалидный статус: {status}"])
   335|                    continue
   336|                
   337|                # Data for "Все лиды" (with employee name and sync date)
   338|                date_val = row[0] if len(row) > 0 else datetime.now().strftime("%Y-%m-%d")
   339|                vacancy = row[1] if len(row) > 1 else ""
   340|                city = row[2] if len(row) > 2 else ""
   341|                age = row[5] if len(row) > 5 else ""
   342|                notes = row[7] if len(row) > 7 else ""
   343|                sync_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
   344|                
   345|                all_leads_data.append([date_val, vacancy, city, fio, phone, age, status, notes, emp["name"], sync_date])
   346|                emp_counts[emp["name"]] += 1
   347|                
   348|                # Stats accumulation
   349|                s = emp_stats[emp["name"]]
   350|                s["total"] += 1
   351|                if is_nd(status): s["nd"] += 1
   352|                if is_podpisan(status): s["podp"] += 1
   353|                if is_komissovan(status): s["komiss"] += 1
   354|                try:
   355|                    age_int = int(age)
   356|                    if age_int > 0: s["ages"].append(age_int)
   357|                except: pass
   358|                
   359|                # Data for "Active"
   360|                if is_active_status(status):
   361|                    active_data.append([date_val, emp["name"], fio, phone, vacancy, city, status, "", "", notes])
   362|                
   363|                # History tracking (current snapshot)
   364|                history_data.append([sync_date, emp["name"], fio, phone, "-", status, "System Sync"])
   365|            
   366|        except Exception as e:
   367|            errors.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), emp["name"], "-", f"Ошибка доступа: {e}"])
   368|    
   369|    # Batch write
   370|    if all_leads_data:
   371|        all_leads_sheet.append_rows(all_leads_data)
   372|        CONSOLE.print(f"[green]Все лиды: {len(all_leads_data)} записей[/green]")
   373|    
   374|    if active_data:
   375|        active_sheet.append_rows(active_data)
   376|        CONSOLE.print(f"[green]Активные: {len(active_data)} записей[/green]")
   377|    
   378|    if history_data:
   379|        history_sheet.append_rows(history_data)
   380|    
   381|    # Update employee sheet
   382|    if emp_sheet:
   383|        for emp in EMPLOYEES:
   384|            name = emp["name"]
   385|            count = emp_counts[name]
   386|            s = emp_stats[name]
   387|            # Find row in employee sheet (skip header)
   388|            try:
   389|                emp_rows = emp_sheet.get_all_values()[1:]
   390|                for idx, row in enumerate(emp_rows, start=2):
   391|                    if row[0] == name:
   392|                        # Update count (column 6)
   393|                        emp_sheet.update_cell(idx, 6, count)
   394|                        # Update % (column 7)
   395|                        if s["total"] > 0:
   396|                            pct = ((s["total"] - s["nd"]) / s["total"]) * 100
   397|                            cell = emp_sheet.cell(idx, 7)
   398|                            cell.value = f"{pct:.1f}%"
   399|                            # Color coding (from v8.0)
   400|                            if pct > 70:
   401|                                cell.format({"textFormat": {"foregroundColor": THEME["accent"]}})
   402|                            elif pct > 40:
   403|                                cell.format({"textFormat": {"foregroundColor": THEME["warning"]}})
   404|                            else:
   405|                                cell.format({"textFormat": {"foregroundColor": THEME["accentAlt"]}})
   406|                        break
   407|            except Exception as e:
   408|                CONSOLE.print(f"[red]Ошибка обновления сотрудника {name}: {e}[/red]")
   409|    
   410|    # Log errors
   411|    if errors:
   412|        log_sheet.append_rows(errors)
   413|        CONSOLE.print(f"[yellow]Ошибок записано: {len(errors)}[/yellow]")
   414|    
   415|    # Update stats
   416|    update_stats(client, emp_stats)
   417|    
   418|    # Highlight duplicates
   419|    highlight_duplicates(client, all_leads_sheet)
   420|    
   421|    CONSOLE.print(f"[bold green]Синхронизация v8.0 завершена. Лидов: {len(all_leads_data)}[/bold green]")
   422|
   423|# ─── STATISTICS WITH SPARKLINES ────────────────────────────────────────────
   424|
   425|def update_stats(client: gspread.Client, emp_stats: dict) -> None:
   426|    """
   427|    Calculate statistics with sparklines (ASCII diagrams).
   428|    From v8.0: updateStats with SPARKLINE function.
   429|    """
   430|    CONSOLE.print(Panel.fit("[STATS] РАСЧЁТ СТАТИСТИКИ", style=THEME["accent"]))
   431|    admin = open_admin(client)
   432|    sheet = admin.worksheet("[STATS] Статистика")
   433|    
   434|    # Clear old data (keep header)
   435|    if sheet.row_count > 1:
   436|        sheet.delete_rows(2, sheet.row_count - 1)
   437|    
   438|    rows = []
   439|    sparkline_data = []
   440|    
   441|    for emp in EMPLOYEES:
   442|        name = emp["name"]
   443|        s = emp_stats.get(name, {"total": 0, "nd": 0, "podp": 0, "komiss": 0, "ages": []})
   444|        if s["total"] == 0:
   445|            continue
   446|        
   447|        dozvon = s["total"] - s["nd"]
   448|        pct = (dozvon / s["total"]) * 100 if s["total"] > 0 else 0
   449|        avg_age = sum(s["ages"]) / len(s["ages"]) if s["ages"] else 0
   450|        
   451|        rows.append([
   452|            "Сейчас", name, s["total"], dozvon, s["nd"],
   453|            s["podp"], s["komiss"], f"{pct:.2f}%", f"{avg_age:.1f}"
   454|        ])
   455|        sparkline_data.append([s["total"], dozvon, s["nd"], s["podp"]])
   456|    
   457|    if rows:
   458|        sheet.append_rows(rows)
   459|        # Add ASCII sparklines to column "График" (J)
   460|        for i, sd in enumerate(sparkline_data, start=2):
   461|            spark = sparkline(sd, width=10)
   462|            sheet.update_cell(i, len(STATS_HEADERS), spark)
   463|        
   464|        # Color coding for % column (from v8.0)
   465|        for i, row in enumerate(rows, start=2):
   466|            pct_val = float(row[7].replace("%", ""))
   467|            cell = sheet.cell(i, 8)  # % column
   468|            if pct_val > 70:
   469|                cell.format({"textFormat": {"foregroundColor": THEME["accent"]}})
   470|            elif pct_val > 40:
   471|                cell.format({"textFormat": {"foregroundColor": THEME["warning"]}})
   472|            else:
   473|                cell.format({"textFormat": {"foregroundColor": THEME["accentAlt"]}})
   474|        
   475|        CONSOLE.print(f"[green]Статистика обновлена для {len(rows)} сотрудников[/green]")
   476|    
   477|    CONSOLE.print("[bold green]Расчёт завершён[/bold green]")
   478|
   479|# ─── SHUFFLE LOGIC (Auto-move old ND) ─────────────────────────────────────
   480|
   481|def shuffle_old_nd(client: gspread.Client) -> None:
   482|    """
   483|    Shuffle old ND leads to employee with least leads.
   484|    From v8.0: shuffleOldND + getEmployeeWithLeastLeads.
   485|    """
   486|    CONSOLE.print(Panel.fit("[ROAD] ШАТТЛ НД (РАЗГОН)", style=THEME["accent"]))
   487|    admin = open_admin(client)
   488|    sheet = admin.worksheet("Все лиды")
   489|    
   490|    if sheet.row_count <= 1:
   491|        CONSOLE.print("[yellow]Нет данных для шаттла[/yellow]")
   492|        return
   493|    
   494|    data = sheet.get_all_values()[1:]  # Skip header
   495|    today = datetime.now()
   496|    moved = 0
   497|    
   498|    # Get employee lead counts
   499|    emp_counts = {emp["name"]: 0 for emp in EMPLOYEES}
   500|    for row in data:
   501|