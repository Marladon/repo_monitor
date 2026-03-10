# repo-monitor

Скрипт для мониторинга репозиториев компании. Читает список репозиториев из Google Sheets и выполняет две проверки:

1. **dev-ветки** — выводит репозитории где dev-веток больше одной
2. **Изменения** — выводит репозитории и ветки в которые были запушены коммиты с момента последнего запуска

## Запуск

```
pip install openpyxl
python repo_monitor.py
```

При первом запуске блок изменений сохраняет текущее время как точку отсчёта. Со второго запуска начинает показывать изменения.

## Конфиг

В начале `repo_monitor.py`:

```python
SHEET_ID      = "..."   # ID Google таблицы (из URL)
GITLAB_TOKEN  = "..."   # gitlab.ximc.ru → Settings → Access Tokens (read_api)
GITHUB_TOKEN  = "..."   # GitHub → Settings → Developer Settings → PAT (repo)
```

## Требования

- Python 3.6+
- `openpyxl`
- Google таблица открыта для чтения по ссылке
- Токены GitLab и GitHub
