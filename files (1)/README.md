# Django Table Processor — Плагин для Р7 Офис

## Архитектура решения

```
┌─────────────────────────────┐        ┌──────────────────────────────┐
│       Р7 Офис               │        │       Django сервер           │
│                             │        │                              │
│  ┌─────────────────────┐   │  JSON  │  ┌────────────────────────┐  │
│  │  Плагин (JS)        │──────POST──►  │  /api/process/         │  │
│  │  - читает ячейки    │   │        │  │  - парсит данные       │  │
│  │  - отправляет data  │   │        │  │  - вызывает макрос     │  │
│  │  - принимает ответ  │◄──────────│  │  - возвращает JSON     │  │
│  │  - пишет в ячейки   │   │        │  └────────────────────────┘  │
│  └─────────────────────┘   │        └──────────────────────────────┘
└─────────────────────────────┘
```

---

## Способы реализации (сравнение)

### Способ 1 — JSON API (этот проект) ✅ Рекомендуется
Плагин читает данные ячеек → передаёт как `{"data": [[...], ...]}` → сервер обрабатывает → возвращает новый массив.

**Плюсы:** Простой протокол, легко отлаживать, не зависит от формата файла.
**Минусы:** Нет метаданных ячеек (формулы, стили, объединения).

### Способ 2 — Передача файла (XLSX)
Плагин сохраняет файл через API Р7, читает как binary → отправляет multipart/form-data → сервер обрабатывает через openpyxl → возвращает новый XLSX → плагин открывает.

**Плюсы:** Сохраняются все стили, формулы, форматирование.
**Минусы:** Сложнее в плагине (бинарные данные), медленнее.

### Способ 3 — WebSocket
Для долгих операций: открываем WS-соединение, сервер шлёт прогресс в реальном времени.

**Плюсы:** Стриминг результатов, индикатор прогресса.
**Минусы:** Сложнее инфраструктура (нужен channels или asyncio).

---

## Структура файлов плагина

```
r7-plugin/
├── config.json        ← метаданные плагина (обязательно)
├── index.html         ← UI панель (открывается в боковой панели)
├── pluginCode.js      ← вся логика плагина
└── resources/
    └── img/
        ├── icon.png   ← иконка 40×40
        └── icon@2x.png  ← иконка 80×80
```

---

## Этап 1 — Структура плагина Р7 Офис

Р7 Офис использует ту же платформу что и OnlyOffice.
Плагин = папка с `config.json` + HTML/JS файлами.

### config.json — ключевые поля:
```json
{
  "guid": "asc.{уникальный-uuid}",   ← должен быть уникальным
  "EditorsSupport": ["cell"],         ← "cell" = таблицы, "word" = текст
  "isVisual": true,                   ← показывать UI панель
  "isModal": true,                    ← открывать как модальное окно
  "size": [450, 380]                  ← размер панели в пикселях
}
```

### Установка плагина:
1. Сервер Р7: скопировать папку в `sdkjs-plugins/` на сервере
2. Десктоп Р7: скопировать в `~/r7-office/sdkjs-plugins/`
3. Веб Р7: Плагины → Добавить плагин → указать путь к config.json

---

## Этап 2 — JavaScript API для работы с ячейками

### Читаем данные через callCommand:
```javascript
window.Asc.plugin.callCommand(function () {
  // Этот код выполняется ВНУТРИ контекста редактора
  const sheet = Api.GetActiveSheet();
  const range = sheet.GetSelection();        // или GetUsedRange()
  const rows  = range.GetRowsCount();
  const cols  = range.GetColumnsCount();

  const data = [];
  for (let r = 0; r < rows; r++) {
    const row = [];
    for (let c = 0; c < cols; c++) {
      row.push(range.GetCell(r, c).GetText());
    }
    data.push(row);
  }
  return data;  // ← передаётся обратно в callback
}, false, function (result) {
  // result — данные из таблицы
  console.log(result);
});
```

### Важное ограничение!
Код внутри `callCommand` **сериализуется и выполняется в другом контексте**.
Нельзя обращаться к переменным плагина напрямую — только через `return`.
Для передачи данных В команду используй `window._myData = ...` до вызова.

### Пишем данные обратно:
```javascript
window._dataToWrite = { data: [["A","B"],["1","2"]], startRow: 0 };

window.Asc.plugin.callCommand(function () {
  const info = window._dataToWrite;
  const sheet = Api.GetActiveSheet();

  info.data.forEach((row, r) => {
    row.forEach((val, c) => {
      sheet.GetRangeByNumber(info.startRow + r, c).SetValue(val);
    });
  });
}, false, function () {
  console.log("Записано!");
});
```

---

## Этап 3 — Django API

### Добавляем endpoint в urls.py:
```python
from django.urls import path
from .views import process_table_api

urlpatterns = [
    path('api/process/', process_table_api, name='process_table'),
]
```

### Настраиваем CORS (обязательно!):
```bash
pip install django-cors-headers
```
```python
# settings.py
INSTALLED_APPS += ['corsheaders']
MIDDLEWARE.insert(0, 'corsheaders.middleware.CorsMiddleware')

# Разрешить запросы от плагина:
CORS_ALLOW_ALL_ORIGINS = True
# Или конкретный origin:
# CORS_ALLOWED_ORIGINS = ["https://your-r7-server.com"]
```

### Формат запроса (POST JSON):
```json
{
    "data": [["Товар","Цена"],["Яблоко","50"],["Банан","30"]],
    "rows": 3,
    "cols": 2,
    "macro": "normalize",
    "params": {"has_header": true},
    "sheet_name": "Лист1"
}
```

### Формат ответа:
```json
{
    "success": true,
    "data": [["Товар","Цена"],["Яблоко","1.0"],["Банан","0.0"]],
    "start_row": 0,
    "start_col": 0,
    "stats": {"processed_rows": 2}
}
```

---

## Этап 4 — Добавление своих макросов

В `django_views.py` добавь функцию и зарегистрируй в `MACROS`:

```python
def macro_my_custom(data: list[list], params: dict) -> dict:
    # data — список строк, каждая строка — список значений ячеек (строки)
    result = []
    for row in data:
        new_row = []
        for cell in row:
            # твоя логика обработки
            new_row.append(cell.upper())
        result.append(new_row)

    return {
        "data": result,
        "stats": {"processed_rows": len(result)},
        # Опционально — куда записать (иначе перезаписывает исходное):
        # "start_row": 0,
        # "start_col": len(data[0]) + 1,  # правее исходных данных
    }

MACROS["my_custom"] = macro_my_custom
```

Затем добавь опцию в `index.html`:
```html
<option value="my_custom">my_custom — моя обработка</option>
```

---

## Этап 5 — Установка и запуск

### 1. Сервер Django:
```bash
pip install django djangorestframework django-cors-headers
# Скопировать django_views.py в приложение
# Добавить путь в urls.py
python manage.py runserver 0.0.0.0:8000
```

### 2. Плагин Р7:
```
Создать папку r7-plugin/ с файлами:
  config.json, index.html, pluginCode.js, resources/img/icon.png

В index.html заменить SERVER_URL на адрес своего сервера.
```

### 3. Подключение плагина:
- **Веб-версия**: Плагины → Управление плагинами → Добавить
- **Десктоп**: скопировать папку в директорию плагинов Р7

---

## Отладка

### Консоль плагина:
В Р7 Офис: F12 (или инструменты разработчика) → Console

### Проверка API без плагина (curl):
```bash
curl -X POST https://your-server.com/api/process/ \
  -H "Content-Type: application/json" \
  -d '{"data":[["1","2"],["3","4"]],"macro":"default","params":{}}'
```

### Частые ошибки:
- **CORS error** → добавить django-cors-headers (см. Этап 3)
- **CSRF error** → добавить @csrf_exempt на view (уже есть в коде)
- **callCommand не возвращает данные** → убедись что return стоит ВНУТРИ function
- **window._pluginWriteData is undefined** → присвоить ДО callCommand
