#!/usr/bin/env python3
# callmonstr v3.0 — Recruiting CRM Admin Script
# Dark theme: #1a1a1a background, #00ff88 accent, Comfortaa font (via rich)
# Features: sync, shuffle, backup, stats, sparklines, history
# Status "🎗️ Комиссован" counts as dozvon

import os
import json
import csv
import random
import shutil
from datetime import datetime
from pathlib import Path

try:
    from rich.console import Console
    from rich.table import Table
    from rich.panel import Panel
    from rich.text import Text
    from rich import box
    from rich.spinner import Spinner
except ImportError:
    print("Установи rich: pip install rich")
    exit(1)

# ── Config ────────────────────────────────────────────────────────────────────
HOME = Path.home() / "callmonstr"
DATA_DIR = HOME / "data"
BACKUP_DIR = HOME / "backups"
HISTORY_FILE = HOME / "history.json"
CONFIG_FILE = HOME / "config.json"

# Valid statuses (Комиссован counts as dozvon)
VALID_STATUSES = {
    "⚪ Новый", "🔴 НД", "🤔 ДУМ", "🟡 ПЕРЕЗВОНИТЬ",
    "💬 СВЯЗЬ МЕССЕНДЖЕР", "✅ ПОДПИСАН", "✅ Подписан",
    "⚫ ОТКАЗ", "❌ Отказ", "📝 Заявка", "🎫 Ожидает билеты",
    "🚗 Ожидает выезда", "🚀 В пути", "🏛️ В военкомате",
    "🔍 На проверке", "🎗️ Комиссован"
}
DOZVON_STATUSES = {"✅ ПОДПИСАН", "✅ Подписан", "🎗️ Комиссован"}

# ── Console setup (dark theme) ────────────────────────────────────────────────
console = Console(
    style="bold #00ff88 on #1a1a1a",
    highlight=True,
)

def header(text: str):
    console.print(Panel.fit(f"[bold #00ff88]{text}[/]", box=box.DOUBLE, border_style="#00ff88"))

def load_config():
    if CONFIG_FILE.exists():
        return json.loads(CONFIG_FILE.read_text())
    return {"table_url": "", "last_sync": None}

def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2, ensure_ascii=False))

def load_data():
    """Load CRM data from CSV or JSON in DATA_DIR."""
    files = list(DATA_DIR.glob("*.csv")) + list(DATA_DIR.glob("*.json"))
    if not files:
        return []
    latest = max(files, key=lambda f: f.stat().st_mtime)
    if latest.suffix == ".csv":
        with open(latest, newline="", encoding="utf-8") as f:
            return list(csv.DictReader(f))
    else:
        return json.loads(latest.read_text(encoding="utf-8"))

def save_data(data, filename=None):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not filename:
        filename = f"crm_{datetime.now():%Y%m%d_%H%M%S}.csv"
    path = DATA_DIR / filename
    if not data:
        return path
    writer = csv.DictWriter(path, fieldnames=data[0].keys())
    writer.writeheader()
    writer.writerows(data)
    return path

def display_table(data, title="CRM Data"):
    if not data:
        console.print("[dim]Нет данных[/]")
        return
    table = Table(title=title, box=box.MINIMAL_DOUBLE_HEAD, border_style="#00ff88")
    for col in data[0].keys():
        table.add_column(col, style="#00ff88")
    for row in data[:30]:  # show first 30 rows
        table.add_row(*[row.get(c, "") for c in data[0].keys()])
    console.print(table)
    if len(data) > 30:
        console.print(f"[dim]... ещё {len(data)-30} строк[/]")

def cmd_sync(args):
    """Sync with remote table (stub — replace with real API call)."""
    cfg = load_config()
    header("Синхронизация")
    if not cfg.get("table_url"):
        console.print("[yellow]Не задан URL таблицы. Используй /set_table <url>[/]")
        return
    console.print(f"Синхронизация с {cfg['table_url']} ...")
    # TODO: implement real sync (Google Sheets API / CSV export)
    # For now, create sample data if none exists
    if not list(DATA_DIR.glob("*")):
        sample = [
            {"id": "1", "name": "Иван", "status": "⚪ Новый", "phone": "+7911..."},
            {"id": "2", "name": "Мария", "status": "✅ Подписан", "phone": "+7922..."},
            {"id": "3", "name": "Петр", "status": "🎗️ Комиссован", "phone": "+7933..."},
        ]
        save_data(sample, "crm_current.csv")
        console.print("[green]Создан пример данных[/]")
    cfg["last_sync"] = datetime.now().isoformat()
    save_config(cfg)
    console.print("[bold green]Готово[/]")

def cmd_shuffle(args):
    """Shuffle rows randomly."""
    data = load_data()
    if not data:
        console.print("[red]Нет данных для шаффла[/]")
        return
    random.shuffle(data)
    save_data(data, "crm_shuffled.csv")
    header("Шаффл завершён")
    console.print(f"Перемешано {len(data)} строк → crm_shuffled.csv")

def cmd_backup(args):
    """Backup current data."""
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    data = load_data()
    if not data:
        console.print("[red]Нет данных для бэкапа[/]")
        return
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"backup_{ts}.csv"
    save_data(data, str(backup_path.name))
    shutil.copy(DATA_DIR / "crm_current.csv", backup_path) if (DATA_DIR / "crm_current.csv").exists() else None
    # also backup history
    if HISTORY_FILE.exists():
        shutil.copy(HISTORY_FILE, BACKUP_DIR / f"history_{ts}.json")
    header("Бэкап создан")
    console.print(f"Сохранено в {backup_path}")

def cmd_stats(args):
    """Show statistics (dozvon counts Комиссован as подписан)."""
    data = load_data()
    if not data:
        console.print("[red]Нет данных[/]")
        return
    stats = {}
    dozvon = 0
    for row in data:
        st = row.get("status", "Неизвестен")
        stats[st] = stats.get(st, 0) + 1
        if st in DOZVON_STATUSES:
            dozvon += 1
    header("Статистика")
    table = Table(box=box.SIMPLE)
    table.add_column("Статус", style="#00ff88")
    table.add_column("Кол-во", justify="right")
    for st, cnt in sorted(stats.items(), key=lambda x: -x[1]):
        table.add_row(st, str(cnt))
    console.print(table)
    console.print(f"\n[bold #00ff88]Дозвон (вкл. Комиссован): {dozvon} из {len(data)}[/]")

def sparkline(numbers, width=20):
    """ASCII sparkline."""
    if not numbers:
        return "─" * width
    mn, mx = min(numbers), max(numbers)
    if mx == mn:
        return "█" * width
    bars = "▁▂▃▄▅▆▇█"
    result = ""
    for n in numbers:
        idx = int((n - mn) / (mx - mn) * (len(bars) - 1))
        result += bars[idx]
    return result

def cmd_sparklines(args):
    """Show sparklines for recent activity."""
    # stub: generate random data for demo
    import random
    data = [random.randint(0, 20) for _ in range(14)]
    header("Спарклайны (14 дней)")
    console.print(f"[dim]Дозвоны:[/] {sparkline(data)}")
    console.print(f"[dim]Значения:[/] {data}")

def cmd_history(args):
    """Show change history."""
    if not HISTORY_FILE.exists():
        console.print("[dim]История пуста[/]")
        return
    history = json.loads(HISTORY_FILE.read_text())
    header("История изменений")
    for entry in history[-10:]:
        ts = entry.get("time", "?")
        action = entry.get("action", "?")
        console.print(f"[dim]{ts}[/] {action}")

def log_history(action: str):
    history = []
    if HISTORY_FILE.exists():
        history = json.loads(HISTORY_FILE.read_text())
    history.append({"time": datetime.now().isoformat(), "action": action})
    HISTORY_FILE.write_text(json.dumps(history[-100:], ensure_ascii=False, indent=2))

def cmd_set_table(args):
    """Set remote table URL."""
    if not args:
        console.print("[red]Укажи URL таблицы[/]")
        return
    cfg = load_config()
    cfg["table_url"] = args[0]
    save_config(cfg)
    console.print(f"[green]Таблица установлена: {args[0]}[/]")
    log_history(f"set_table {args[0]}")

def cmd_help(args):
    header("Доступные команды")
    cmds = [
        ("sync", "Синхронизировать с таблицей"),
        ("shuffle", "Перемешать строки"),
        ("backup", "Создать бэкап"),
        ("stats", "Показать статистику"),
        ("sparklines", "Спарклайны активности"),
        ("history", "История изменений"),
        ("set_table <url>", "Установить URL таблицы"),
        ("list", "Показать данные"),
        ("help", "Эта справка"),
    ]
    for cmd, desc in cmds:
        console.print(f"  [bold #00ff88]{cmd}[/] — {desc}")

def cmd_list(args):
    data = load_data()
    display_table(data, title="Текущие данные")

COMMANDS = {
    "sync": cmd_sync,
    "shuffle": cmd_shuffle,
    "backup": cmd_backup,
    "stats": cmd_stats,
    "sparklines": cmd_sparklines,
    "history": cmd_history,
    "set_table": cmd_set_table,
    "help": cmd_help,
    "list": cmd_list,
}

def main():
    console.print(Panel.fit(
        "[bold #00ff88]callmonstr v3.0[/]\n[dim]Recruiting CRM Admin[/]",
        box=box.DOUBLE, border_style="#00ff88"
    ))
    # simple REPL
    while True:
        try:
            inp = console.input("[bold #00ff88]callmonstr>[/] ").strip()
        except (EOFError, KeyboardInterrupt):
            console.print("\n[dim]Пока[/]")
            break
        if not inp:
            continue
        parts = inp.split(maxsplit=1)
        cmd = parts[0].lower()
        args = [parts[1]] if len(parts) > 1 else []
        if cmd in ("exit", "quit", "q"):
            console.print("[dim]Пока[/]")
            break
        if cmd in COMMANDS:
            try:
                COMMANDS[cmd](args)
                log_history(f"{cmd} {' '.join(args)}")
            except Exception as e:
                console.print(f"[red]Ошибка: {e}[/]")
        else:
            console.print(f"[yellow]Неизвестная команда: {cmd}. Введи 'help'[/]")

if __name__ == "__main__":
    main()
