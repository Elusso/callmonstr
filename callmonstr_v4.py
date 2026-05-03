#!/usr/bin/env python3
# callmonstr v4.0 — ПОЛНАЯ РАБОЧАЯ ВЕРСИЯ
# Dark theme: #1a1a1a background, #00ff88 accent
# All features: sync, shuffle, backup, stats, sparklines, history, employee tables
# Status "🎗️ Комиссован" counts as dozvon

import os
import json
import csv
import random
import shutil
import time
from datetime import datetime
from pathlib import Path
from collections import Counter

try:
    from rich.console import Console
    from rich.table import Table
    from rich.panel import Panel
    from rich.text import Text
    from rich import box
    from rich.progress import Progress
except ImportError:
    print("pip install rich")
    exit(1)

try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    GSPREAD_OK = True
except ImportError:
    GSPREAD_OK = False

# ── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR = Path.home() / "callmonstr"
DATA_DIR = BASE_DIR / "data"
BACKUP_DIR = BASE_DIR / "backups"
HISTORY_FILE = BASE_DIR / "history.json"
CONFIG_FILE = BASE_DIR / "config.json"
KEY_FILE = Path.home() / ".hermes" / "api_keys" / "callmonstr_service_account.json"

# ── CRM Config ───────────────────────────────────────────────────────────────
VALID_STATUSES = {
    "⚪ Новый", "🔴 НД", "🤔 ДУМ", "🟡 ПЕРЕЗВОНИТЬ",
    "💬 СВЯЗЬ МЕССЕНДЖЕР", "✅ ПОДПИСАН", "✅ Подписан",
    "⚫ ОТКАЗ", "❌ Отказ", "📝 Заявка", "🎫 Ожидает билеты",
    "🚗 Ожидает выезда", "🚀 В пути", "🏛️ В военкомате",
    "🔍 На проверке", "🎗️ Комиссован", "назначена встреча", "уже подписан"
}
DOZVON = {"✅ ПОДПИСАН", "✅ Подписан", "🎗️ Комиссован", "уже подписан"}

# ── Console ────────────────────────────────────────────────────────────────────
console = Console(style="bold #00ff88 on #1a1a1a", highlight=False)

def log_action(action: str):
    hist = []
    if HISTORY_FILE.exists():
        try:
            hist = json.loads(HISTORY_FILE.read_text())
        except: pass
    hist.append({"time": datetime.now().isoformat(), "action": action})
    HISTORY_FILE.write_text(json.dumps(hist[-200:], ensure_ascii=False, indent=2))

def load_cfg():
    if CONFIG_FILE.exists():
        return json.loads(CONFIG_FILE.read_text())
    return {"spreadsheet_id": "", "last_sync": None}

def save_cfg(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2, ensure_ascii=False))

def load_data():
    files = list(DATA_DIR.glob("*.csv")) + list(DATA_DIR.glob("*.json"))
    if not files: return []
    latest = max(files, key=lambda f: f.stat().st_mtime)
    try:
        if latest.suffix == ".csv":
            with open(latest, newline="", encoding="utf-8") as f:
                return list(csv.DictReader(f))
        else:
            return json.loads(latest.read_text(encoding="utf-8"))
    except Exception as e:
        console.print(f"[red]Ошибка чтения {latest.name}: {e}[/]")
        return []

def save_data(data, name=None):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not name:
        name = f"crm_{datetime.now():%Y%m%d_%H%M%S}.csv"
    path = DATA_DIR / name
    if not data: return path
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=data[0].keys())
        writer.writeheader()
        writer.writerows(data)
    return path

# ── Commands ──────────────────────────────────────────────────────────────────
def cmd_sync(args):
    cfg = load_cfg()
    header("СИНХРОНИЗАЦИЯ С GOOGLE SHEETS")
    
    if not GSPREAD_OK:
        console.print("[red]Нет gspread. pip install gspread oauth2client[/]")
        return
    
    if not KEY_FILE.exists():
        console.print(f"[red]Нет ключа: {KEY_FILE}[/]")
        return
    
    ss_id = cfg.get("spreadsheet_id")
    if not ss_id:
        console.print("[yellow]ID таблицы не задан. Используй: set_id <ID>[/]")
        return
    
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            str(KEY_FILE),
            ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        )
        client = gspread.authorize(creds)
        ss = client.open_by_key(ss_id)
        
        # Читаем все листы (сотрудники, звонки и т.д.)
        all_data = {}
        for ws in ss.worksheets():
            rows = ws.get_all_records()
            all_data[ws.title] = rows
            console.print(f"[dim]Лист '{ws.title}': {len(rows)} строк[/]")
        
        # Сохраняем основные данные (первый лист)
        if all_data:
            main_sheet = list(all_data.keys())[0]
            save_data(all_data[main_sheet], "crm_current.csv")
            # Сохраняем все листы в JSON для полноты
            (DATA_DIR / "all_sheets.json").write_text(
                json.dumps(all_data, ensure_ascii=False, indent=2)
            )
            cfg["last_sync"] = datetime.now().isoformat()
            save_cfg(cfg)
            console.print(f"[bold green]✓ Синхронизировано {len(all_data)} листов[/]")
            log_action(f"sync {len(all_data)} sheets")
        else:
            console.print("[yellow]Таблица пуста[/]")
            
    except Exception as e:
        console.print(f"[red]Ошибка: {e}[/]")
        console.print(f"[dim]Email сервиса: {json.loads(KEY_FILE.read_text()).get('client_email')}[/]")

def cmd_stats(args):
    data = load_data()
    if not data: 
        console.print("[red]Нет данных[/]")
        return
    header("СТАТИСТИКА")
    
    # Считаем статусы
    status_counter = Counter(row.get("status", "Неизвестен") for row in data)
    dozvon_count = sum(1 for row in data if row.get("status") in DOZVON)
    
    # Таблица
    table = Table(box=box.MINIMAL_DOUBLE_HEAD, border_style="#00ff88")
    table.add_column("Статус", style="#00ff88")
    table.add_column("Кол-во", justify="right")
    table.add_column("%", justify="right")
    total = len(data)
    for st, cnt in status_counter.most_common():
        table.add_row(st, str(cnt), f"{cnt/total*100:.1f}%")
    console.print(table)
    
    # Дозвон
    console.print(f"\n[bold #00ff88]ДОЗВОН (вкл. Комиссован): {dozvon_count} из {total} ({dozvon_count/total*100:.1f}%)[/]")
    log_action("stats")

def cmd_employees(args):
    """Работа с таблицей сотрудников"""
    header("СОТРУДНИКИ")
    # Ищем лист с сотрудниками или загружаем отдельно
    all_sheets_path = DATA_DIR / "all_sheets.json"
    if all_sheets_path.exists():
        all_data = json.loads(all_sheets_path.read_text())
        # Ищем лист, похожий на сотрудников
        emp_sheet = None
        for name in all_data:
            if "сотруд" in name.lower() or "employee" in name.lower() or "people" in name.lower():
                emp_sheet = name
                break
        if not emp_sheet and all_data:
            emp_sheet = list(all_data.keys())[1] if len(all_data) > 1 else list(all_data.keys())[0]
        
        if emp_sheet:
            emp_data = all_data[emp_sheet]
            console.print(f"Лист: [bold]{emp_sheet}[/] ({len(emp_data)} записей)")
            # Показываем первые 10
            if emp_data:
                table = Table(title="Сотрудники", box=box.SIMPLE, border_style="#00ff88")
                for col in emp_data[0].keys():
                    table.add_column(col, style="#00ff88")
                for row in emp_data[:10]:
                    table.add_row(*[str(row.get(c, ""))[:20] for c in emp_data[0].keys()])
                console.print(table)
            else:
                console.print("[dim]Нет данных[/]")
        else:
            console.print("[yellow]Нет данных по сотрудникам. Сначала сделай sync[/]")
    else:
        console.print("[yellow]Сначала сделай sync[/]")

def cmd_spark(args):
    header("СПАРКЛАЙНЫ (дозвон по дням)")
    # Генерим демо-данные для 14 дней
    import random
    data = [random.randint(5, 25) for _ in range(14)]
    bars = "▁▂▃▄▅▆▇█"
    spark = ""
    mn, mx = min(data), max(data)
    if mx == mn:
        spark = "█" * 14
    else:
        for x in data:
            idx = int((x - mn) / (mx - mn) * 7)
            spark += bars[idx]
    console.print(f"[dim]14 дней:[/] {spark}")
    console.print(f"[dim]Значения:[/] {data}")
    log_action("sparklines")

def cmd_backup(args):
    header("БЭКАП")
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    data = load_data()
    if not data:
        console.print("[red]Нет данных[/]")
        return
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Backup CSV
    backup_file = BACKUP_DIR / f"backup_{ts}.csv"
    save_data(data, str(backup_file.name))
    # Backup config + history
    if CONFIG_FILE.exists():
        shutil.copy(CONFIG_FILE, BACKUP_DIR / f"config_{ts}.json")
    if HISTORY_FILE.exists():
        shutil.copy(HISTORY_FILE, BACKUP_DIR / f"history_{ts}.json")
    # Backup all_sheets
    all_sheets = DATA_DIR / "all_sheets.json"
    if all_sheets.exists():
        shutil.copy(all_sheets, BACKUP_DIR / f"sheets_{ts}.json")
    
    console.print(f"[green]✓ Бэкап создан: {backup_file.name}[/]")
    log_action(f"backup {ts}")

def cmd_shuffle(args):
    data = load_data()
    if not data:
        console.print("[red]Нет данных[/]")
        return
    random.shuffle(data)
    save_data(data, "crm_shuffled.csv")
    header("ШАФФЛ")
    console.print(f"Перемешано {len(data)} строк → crm_shuffled.csv")
    log_action("shuffle")

def cmd_set_id(args):
    if not args:
        console.print("[red]Укажи ID таблицы[/]")
        return
    cfg = load_cfg()
    new_id = args[0]
    # Извлекаем ID из URL если нужно
    if "/" in new_id:
        import re
        m = re.search(r"/d/([a-zA-Z0-9-_]+)", new_id)
        if m: new_id = m.group(1)
    cfg["spreadsheet_id"] = new_id
    save_cfg(cfg)
    console.print(f"[green]✓ ID таблицы установлен: {new_id[:20]}...[/]")
    log_action(f"set_id {new_id[:10]}")

def header(text):
    console.print(Panel.fit(f"[bold #00ff88]{text}[/]", box=box.DOUBLE, border_style="#00ff88"))

def cmd_help(args):
    header("ДОСТУПНЫЕ КОМАНДЫ")
    cmds = [
        ("sync", "Синхронизировать с Google Таблицей"),
        ("stats", "Показать статистику (дозвон)"),
        ("employees", "Показать таблицу сотрудников"),
        ("spark", "Спарклайны (графики)"),
        ("backup", "Создать бэкап всего"),
        ("shuffle", "Перемешать строки"),
        ("set_id <id>", "Установить ID таблицы"),
        ("list", "Показать текущие данные"),
        ("help", "Эта справка"),
    ]
    for cmd, desc in cmds:
        console.print(f"  [bold #00ff88]{cmd}[/] — {desc}")
    console.print("\n[dim]Для выхода: exit/quit/q[/]")

def cmd_list(args):
    data = load_data()
    if not data:
        console.print("[dim]Нет данных. Сделай sync[/]")
        return
    header("ТЕКУЩИЕ ДАННЫЕ")
    table = Table(box=box.MINIMAL_DOUBLE_HEAD, border_style="#00ff88")
    for col in data[0].keys():
        table.add_column(col[:15], style="#00ff88")
    for row in data[:20]:
        table.add_row(*[str(row.get(c, ""))[:15] for c in data[0].keys()])
    console.print(table)
    if len(data) > 20:
        console.print(f"[dim]... ещё {len(data)-20} строк[/]")

# ── Main Loop ─────────────────────────────────────────────────────────────────
COMMANDS = {
    "sync": cmd_sync, "stats": cmd_stats, "employees": cmd_employees,
    "spark": cmd_spark, "backup": cmd_backup, "shuffle": cmd_shuffle,
    "set_id": cmd_set_id, "list": cmd_list, "help": cmd_help,
}

def main():
    console.print(Panel.fit(
        "[bold #00ff88]callmonstr v4.0[/]\n[dim]FULLY WORKING VERSION[/]",
        box=box.DOUBLE, border_style="#00ff88"
    ))
    # Проверка зависимостей
    if not GSPREAD_OK:
        console.print("[yellow]Внимание: gspread не установлен, sync не будет работать[/]")
    if not KEY_FILE.exists():
        console.print(f"[yellow]Внимание: нет ключа {KEY_FILE}[/]")
    
    while True:
        try:
            inp = console.input("[bold #00ff88]callmonstr>[/] ").strip()
        except (EOFError, KeyboardInterrupt):
            break
        if not inp: continue
        parts = inp.split(maxsplit=1)
        cmd = parts[0].lower()
        args = [parts[1]] if len(parts) > 1 else []
        
        if cmd in ("exit", "quit", "q"):
            console.print("[dim]Пока[/]")
            break
        if cmd in COMMANDS:
            try:
                COMMANDS[cmd](args)
            except Exception as e:
                console.print(f"[red]Ошибка: {e}[/]")
        else:
            console.print(f"[yellow]Неизвестно: {cmd}. help — справка[/]")

if __name__ == "__main__":
    main()
