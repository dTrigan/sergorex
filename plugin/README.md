# Django Processor — Плагин для Р7 Офис

Интеграция Р7 Офис ↔ Django: данные таблицы отправляются на сервер,
обрабатываются Python-макросом, и результат вставляется обратно в таблицу.

---

## Структура файлов

```
r7_plugin/
└── plugin/
    ├── config.json   ← манифест плагина
    └── index.html    ← UI + логика (JS)

django_api/
└── api_views.py      ← Django views с эндпоинтами
```

---

## Установка плагина в Р7 Офис

### Способ 1 — через интерфейс (рекомендуется)
1. Откройте Р7 Офис Редактор таблиц
2. Плагины → Менеджер плагинов → Добавить плагин
3. Укажите путь к папке `plugin/` или загрузите `config.json`

### Способ 2 — вручную
Скопируйте папку `plugin/` в директорию плагинов Р7 Офис:
- **Windows:** `%appdata%\R7-Office\sdkjs-plugins\`
- **Linux:**   `~/.local/share/r7-office/sdkjs-plugins/`
- **macOS:**   `~/Library/Application Support/R7-Office/sdkjs-plugins/`

---

## Настройка Django-сервера

### 1. Установить зависимости

```bash
pip install django django-cors-headers
```

### 2. Добавить в settings.py

```python
INSTALLED_APPS += ['corsheaders']

MIDDLEWARE = [
    'corsheaders.middleware.CorsMiddleware',  # ПЕРВЫМ в списке!
    ...
]

# Разрешить запросы от плагина (или укажи конкретный origin)
CORS_ALLOW_ALL_ORIGINS = True
```

### 3. Добавить в urls.py

```python
from django.urls import path
from . import api_views

urlpatterns += [
    path('api/process/', api_views.process_table, name='api_process'),
    path('api/ping/',    api_views.ping,           name='api_ping'),
]
```

### 4. Скопировать api_views.py в своё Django-приложение

```bash
cp django_api/api_views.py myapp/api_views.py
```

---

## Формат обмена данными

### Запрос (плагин → сервер)

```json
POST /api/process/
Content-Type: application/json

{
  "sheet":      "Sheet1",
  "range":      "A1:D10",
  "hasHeaders": true,
  "data": [
    ["Имя", "Возраст", "Город", "Оценка"],
    ["Алиса", 28, "Москва", 85],
    ["Борис", 34, "СПб",    92]
  ]
}
```

### Ответ (сервер → плагин)

```json
{
  "data": [
    ["Имя", "Возраст", "Город", "Оценка"],
    ["Алиса", 56, "Москва", 170],
    ["Борис", 68, "СПб",    184],
    ["ИТОГО:", 124, "", 354]
  ],
  "info": "Обработано 2 строки, добавлена итоговая строка"
}
```

---

## Замена макроса своим кодом

Вся логика обработки — в функции `run_macro()` в файле `api_views.py`:

```python
def run_macro(data, has_headers, sheet, range_str):
    # data — 2D список [[строка], [строка], ...]
    # Верни {"data": [[...], ...], "info": "пояснение"}
    
    import pandas as pd
    headers = data[0] if has_headers else None
    rows = data[1:] if has_headers else data
    
    df = pd.DataFrame(rows, columns=headers)
    # ... твоя обработка ...
    df_result = df  # или любой другой DataFrame
    
    result_data = []
    if headers:
        result_data.append(df_result.columns.tolist())
    result_data.extend(df_result.values.tolist())
    
    return {"data": result_data, "info": "Готово"}
```

---

## Режимы вставки результата

| Режим          | Описание                                          |
|----------------|---------------------------------------------------|
| На то же место | Перезаписывает исходный диапазон                  |
| Новый столбец  | Добавляет результат правее последнего столбца     |
| Новый лист     | Создаёт лист «Результат_ЧЧ-ММ-СС» с результатом  |

---

## Проверка соединения

Кнопка «Проверить соединение» делает GET-запрос на `/api/ping/`.
Убедитесь, что Django-сервер запущен и доступен по указанному URL.

```bash
# Тест вручную
curl http://localhost:8000/api/ping/
# → {"status": "ok", "service": "Django Processor"}
```

---

## Частые проблемы

**Ошибка CORS** — добавьте `corsheaders` в Django (см. выше).

**fetch не работает в плагине** — Р7 Офис разрешает `fetch` только на HTTPS
или на `localhost`. Для продакшн-сервера используйте HTTPS.

**Пустые данные** — нажмите «↗ выбор» чтобы автоматически определить
используемый диапазон листа.
