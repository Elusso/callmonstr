# Как дать доступ к Google-таблице

## Вариант 1: Через сервисный аккаунт (рекомендуемый)
1. Создай проект в [Google Cloud Console](https://console.cloud.google.com/)
2. Включи Google Sheets API
3. Создай сервисный аккаунт (IAM & Admin → Service Accounts)
4. Скачай JSON-ключ → сохрани как `service_account.json` в папку `~/callmonstr/`
5. Поделись таблицей с email сервисного аккаунта (формат: `xxx@yyy.iam.gserviceaccount.com`) — права «Редактор»

Скрипт подхватит ключ автоматически.

## Вариант 2: Публичная ссылка (только чтение)
1. В таблице: Файл → Опубликовать в интернете → Выбери «Веб-страница» или «CSV»
2. Скопируй ссылку
3. В скрипте: `python admin_v3.py`
   ```
   callmonstr> set_table <твоя_ссылка>
   callmonstr> sync
   ```

## Вариант 3: Прямой ID таблицы
Скопируй ID из URL: `https://docs.google.com/spreadsheets/d/<ID>/edit`
Используй тот же `set_table` с этим ID.

---
После настройки запусти:
```bash
cd ~/callmonstr
pip install -r requirements.txt
python admin_v3.py
```
