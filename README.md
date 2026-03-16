# repo-monitor

Инструмент для мониторинга репозиториев компании. Читает список репозиториев из Google Sheets и выполняет две проверки:

1. **dev-ветки** — выводит репозитории где dev-веток больше одной
2. **Изменения** — выводит репозитории и ветки в которые были запушены коммиты с момента последнего запуска, а также проверяет было ли слияние между соседними dev-ветками

Доступно в двух вариантах: консольный скрипт и Telegram-бот.

## Файлы

- `repo_monitor.py` — консольный скрипт
- `repo_monitorBot.py` — Telegram-бот с кнопкой запуска

## Запуск

### Консоль
```
pip install openpyxl
python repo_monitor.py
```

### Telegram-бот
```
pip install openpyxl
python repo_monitorBot.py
```

Бот пришлёт сообщение с кнопкой **Запустить проверку**. Перед первым запуском нужно написать боту `/start` в Telegram.

При первом запуске блок изменений сохраняет текущее время как точку отсчёта. Со второго запуска начинает показывать изменения.

## Конфиг

В начале каждого файла:

```python
SHEET_ID      = "..."   # ID Google таблицы (из URL)
GITLAB_TOKEN  = "..."   # gitlab.ximc.ru → Settings → Access Tokens (read_api)
GITHUB_TOKEN  = "..."   # GitHub → Settings → Developer Settings → PAT (repo)
TG_TOKEN      = "..."   # Telegram-бот токен от @BotFather (только для бота)
TG_USER_ID    = ...     # Твой Telegram ID от @userinfobot (только для бота)
```

## Требования

- Python 3.8+
- `openpyxl`
- Google таблица открыта для чтения по ссылке
- Токены GitLab и GitHub