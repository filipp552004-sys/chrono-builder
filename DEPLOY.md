# chrono-builder — развёртывание на Railway

Микросервис собирает карту хронометража (Excel) по данным замеров.
Аналог norma-builder, но отдельный сервис.

## Файлы
- app.py — Flask, эндпоинты /build и /health
- builder.py — генератор Excel (openpyxl, живые формулы)
- requirements.txt, Procfile

## Развёртывание
1. Создать новый сервис на Railway (New Project → Deploy from repo/folder), загрузить эти 4 файла.
2. В Variables добавить: CHRONO_BUILDER_API_KEY = <придумать длинный ключ>
   (PORT Railway задаёт сам.)
3. Дождаться билда. Проверить: GET https://<домен>/health → {"status":"ok"}.
4. Скопировать публичный URL сервиса.

## Секреты в Supabase (для process-chrono, шаг 3)
- CHRONO_BUILDER_URL = https://<домен>
- CHRONO_BUILDER_API_KEY = <тот же ключ>

## Контракт
POST /build
  Заголовки: X-API-Key: <ключ>, Content-Type: application/json
  Тело (пример):
  {
    "operation_name": "Оформление приказа о приёме",
    "position_name": "...", "department_name": "...", "unit_of_work": "документ",
    "work_type": "manual",            // machine | machine_manual | manual | observation
    "observer_name": "...", "obs_date": "12.03.2026",
    "elements": [
      {"name": "Действие 1", "measurements": [{"min": 2.1}, {"min": 2.3}, {"min": 3.5, "excluded": true}]},
      {"name": "Действие 2", "measurements": [{"min": 1.5}, {"min": 1.6}]}
    ],
    "servicing_percent": 6, "rest_percent": 8,
    "prep_final_minutes": 10, "batch_size": 20,
    "norms_source": "..."
  }
  Ответ: 200 + файл .xlsx (bytes). Ошибки: 401 (ключ), 400 (данные), 500 (сборка).
