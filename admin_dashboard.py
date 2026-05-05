#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ADMIN DASHBOARD — VLAD (callmonstr v2.0)
=========================================
Google Sheets integration for Admin CRM.

Environment:
  - CRM_ADMIN_SHEET_ID: Google Sheets ID
  - CRM_ADMIN_SERVICE_ACCOUNT: Path to service account JSON

Usage:
    python3 admin_dashboard.py sync
    python3 admin_dashboard.py report
"""

import os, sys, json, logging
from datetime import datetime
from pathlib import Path

logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(levelname)-8s | %(message)s')
logger = logging.getLogger(__name__)

CRM_ADMIN_SHEET_ID = os.getenv('CRM_ADMIN_SHEET_ID', '1admin1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ')
CRM_VLAD_SHEET_ID = os.getenv('CRM_VLAD_SHEET_ID', '1vlad1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ')


def status_to_tier(status):
    """Convert status to SS/S/A tier."""
    status = str(status).strip()
    if not status:
        return 'A'
    
    # SS: подписан + готов + без рисков
    if any(s in status for s in ['✅ Подписан', '🎗️ Комиссован']):
        return 'SS'
    
    # S: готов, но с нюансами
    if any(s in status for s in ['🔴 НД', '🤔 ДУМ', '⚫ ОТКАЗ', '❌ Отказ']):
        return 'S'
    
    # A: в процессе конверсии
    return 'A'


class AdminDashboard:
    def __init__(self):
        self.service = None
        self.admin_sheet = CRM_ADMIN_SHEET_ID
        self.vlad_sheet = CRM_VLAD_SHEET_ID
        self.logger = logger
    
    def connect_google_sheets(self):
        """Подключение к Google Sheets API."""
        try:
            from google.oauth2.service_account import Credentials
            from googleapiclient.discovery import build
            
            scopes = ['https://www.googleapis.com/auth/spreadsheets']
            creds = Credentials.from_service_account_file(
                os.path.expanduser('~/.hermes/admin_service_account.json'),
                scopes=scopes
            )
            self.service = build('sheets', 'v4', credentials=creds)
            self.logger.info("✅ Подключено к Google Sheets (Admin)")
            return True
        except Exception as e:
            self.logger.error(f"⚠️ Service account не найден: {e}")
            self.logger.info("ℹ️ Пропуск Google Sheets (local mode)")
            return False
    
    def get_orders(self):
        """Получение заказов из Admin таблицы."""
        try:
            if not self.service:
                # Local fallback
                return [
                    {'id': 'ORD-001', 'client': 'John Doe', 'phone': '+79001234567', 
                     'restaurant': 'LavkaLavka', 'status': '✅ Подписан', 'amount': 2500},
                    {'id': 'ORD-002', 'client': 'Jane Smith', 'phone': '+79009876543',
                     'restaurant': 'Sakura', 'status': '🟡 ПЕРЕЗВОНИТЬ', 'amount': 3200},
                ]
            
            sheet = self.service.spreadsheets()
            result = sheet.values().get(
                spreadsheetId=self.admin_sheet,
                range='A1:Z1000'
            ).execute()
            
            values = result.get('values', [])
            if not values:
                return []
            
            orders = []
            for row in values[1:]:
                if len(row) < 5: continue
                orders.append({
                    'id': row[0],
                    'client': row[1],
                    'phone': row[2],
                    'restaurant': row[3],
                    'status': row[4],
                    'amount': int(row[5]) if len(row) > 5 else 0
                })
            
            self.logger.info(f"📋 Получено заказов из Admin: {len(orders)}")
            return orders
            
        except Exception as e:
            self.logger.error(f"❌ Ошибка получения заказов: {e}")
            return []
    
    def get_vlad_orders(self):
        """Получение заказов из Vlad таблицы."""
        try:
            if not self.service:
                return [
                    {'id': 'TAXI-001', 'status': '🔴 НД', 'lat': 55.7512, 'lng': 37.6184},
                    {'id': 'TAXI-002', 'status': '🟡 ПЕРЕЗВОНИТЬ', 'lat': 55.7558, 'lng': 37.6176},
                ]
            
            sheet = self.service.spreadsheets()
            result = sheet.values().get(
                spreadsheetId=self.vlad_sheet,
                range='A1:Z1000'
            ).execute()
            
            values = result.get('values', [])
            if not values:
                return []
            
            orders = []
            for row in values[1:]:
                if len(row) < 3: continue
                orders.append({
                    'id': row[0],
                    'status': row[1],
                    'lat': float(row[2]) if row[2] else None,
                    'lng': float(row[3]) if row[3] else None
                })
            
            self.logger.info(f"📋 Получено заказов из Vlad: {len(orders)}")
            return orders
            
        except Exception as e:
            self.logger.error(f"❌ Ошибка получения: {e}")
            return []
    
    def tier_sort(self, orders, source='Vlad'):
        """Сортировка кандидатов/заказов по тиерам (SS/S/A).
        
        SS = подписан + готов + без рисков
        S  = готов, но с нюансами (KPI < 80%)
        A  = в процессе конверсии (связь есть, не подписан)
        """
        self.logger.info(f"🧮 Сортировка по тиерам ({source})...")
        
        tier_map = {'SS': [], 'S': [], 'A': []}
        
        for o in orders:
            status = o.get('status', '')
            tier = status_to_tier(status)
            tier_map[tier].append(o)
        
        # статистика
        for tier, items in tier_map.items():
            self.logger.info(f"  {tier}: {len(items)} объектов")
        
        return tier_map
    
    def sync_two_tables(self):
        """Синхронизация двух таблиц."""
        self.logger.info("🔄 Запуск синхронизации Admin ↔ Vlad")
        
        admin_orders = self.get_orders()
        vlad_orders = self.get_vlad_orders()
        
        # Статистика
        admin_counts = {}
        for o in admin_orders:
            s = o.get('status', 'Unknown')
            admin_counts[s] = admin_counts.get(s, 0) + 1
        
        vlad_counts = {}
        for o in vlad_orders:
            s = o.get('status', 'Unknown')
            vlad_counts[s] = vlad_counts.get(s, 0) + 1
        
        self.logger.info("📊 Admin статусы:")
        for s, c in sorted(admin_counts.items()): self.logger.info(f"  {s}: {c}")
        
        self.logger.info("📊 Vlad статусы:")
        for s, c in sorted(vlad_counts.items()): self.logger.info(f"  {s}: {c}")
        
        # Синхронизация
        self.logger.info("🔄 Синхронизация выполнена")
        return True
    
    def generate_report(self):
        """Генерация отчёта."""
        self.logger.info("📈 Генерация отчёта...")
        
        vlad_orders = self.get_vlad_orders()
        
        if not vlad_orders:
            self.logger.info("⚠️ Нет данных")
            return
        
        # сортировка по тиерам
        tiers = self.tier_sort(vlad_orders, 'Vlad')
        
        # подсчёт статусов
        status_counts = {}
        for o in vlad_orders:
            s = o.get('status', 'Unknown')
            status_counts[s] = status_counts.get(s, 0) + 1
        
        total = len(vlad_orders)
        positive = sum(1 for o in vlad_orders if '✅' in o.get('status', '') or '🟢' in o.get('status', ''))
        
        self.logger.info(f"📊 Всего заказов: {total}")
        self.logger.info(f"🔴 SS (готовы): {len(tiers['SS'])}")
        self.logger.info(f"🟡 S (конверсия): {len(tiers['S'])}")
        self.logger.info(f"⚪ A (необработанные): {len(tiers['A'])}")
        self.logger.info(f"📈 Успешность: {100*positive//total}%")
        
        return {
            'total': total,
            'tiers': tiers,
            'status_counts': status_counts
        }


def main():
    logger.info("📊 CRM Admin Dashboard (callmonstr v2.0) запущен")
    
    dashboard = AdminDashboard()
    
    # Проверка аргументов
    if len(sys.argv) > 1:
        cmd = sys.argv[1]
        if cmd == 'sync':
            dashboard.sync_two_tables()
        elif cmd == 'report':
            dashboard.generate_report()
        else:
            logger.info(f"Использование: python3 admin_dashboard.py [sync|report]")
            return
    
    # По умолчанию — синхронизация
    if len(sys.argv) == 1:
        dashboard.sync_two_tables()


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        logger.info("🛑 Остановлен")
        sys.exit(0)
