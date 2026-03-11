"""
Django API для обработки данных таблицы из Р7 Офис плагина.

Добавь в urls.py:
    from django.urls import path
    from . import api_views

    urlpatterns += [
        path('api/process/', api_views.process_table, name='api_process'),
        path('api/ping/',    api_views.ping,          name='api_ping'),
    ]

Добавь в settings.py (для CORS, если плагин и сервер на разных портах):
    INSTALLED_APPS += ['corsheaders']
    MIDDLEWARE = ['corsheaders.middleware.CorsMiddleware'] + MIDDLEWARE
    CORS_ALLOW_ALL_ORIGINS = True   # или конкретные origins
"""

import json
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods


# ─────────────────────────────────────────────
#  Твой Python-макрос — вся логика здесь
# ─────────────────────────────────────────────

def run_macro(data: list[list], has_headers: bool, sheet: str, range_str: str) -> dict:
    """
    data        — 2D-список ячеек [[строка], [строка], …]
    has_headers — True если первая строка — заголовки
    sheet       — название листа
    range_str   — диапазон, например "A1:D10"

    Верни dict:
        {
          "data": [[...], ...],   # обработанные данные (2D список)
          "info": "строка-пояснение для пользователя"  # опционально
        }
    """

    # ── Пример 1: добавить итоговую строку с суммами числовых столбцов ──────
    if not data:
        return {"data": data, "info": "Нет данных"}

    headers = data[0] if has_headers else None
    rows    = data[1:] if has_headers else data

    # Определяем числовые столбцы
    result_rows = []
    totals = []

    for row in rows:
        new_row = []
        for cell in row:
            # Пример обработки: умножить все числа на 2
            try:
                new_row.append(float(cell) * 2 if cell != '' else cell)
            except (ValueError, TypeError):
                new_row.append(cell)
        result_rows.append(new_row)

    # Подсчитать суммы числовых столбцов
    col_count = max(len(r) for r in result_rows) if result_rows else 0
    for col_idx in range(col_count):
        values = []
        for row in result_rows:
            val = row[col_idx] if col_idx < len(row) else None
            if isinstance(val, (int, float)):
                values.append(val)
        totals.append(sum(values) if values else '')

    # Собираем результат
    output = []
    if headers:
        output.append(headers)          # заголовки без изменений
    output.extend(result_rows)          # обработанные строки
    output.append(['ИТОГО:'] + totals[1:])  # итоговая строка

    return {
        "data": output,
        "info": f"Обработано {len(result_rows)} строк, добавлена итоговая строка"
    }

    # ── Пример 2: вызов внешнего скрипта ────────────────────────────────────
    # import subprocess, tempfile, pandas as pd
    # df = pd.DataFrame(rows, columns=headers)
    # with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as f:
    #     df.to_csv(f.name, index=False)
    #     result = subprocess.run(['python', 'my_macro.py', f.name], capture_output=True)
    # df_out = pd.read_csv('output.csv')
    # return {"data": [df_out.columns.tolist()] + df_out.values.tolist()}


# ─────────────────────────────────────────────
#  Endpoints
# ─────────────────────────────────────────────

@csrf_exempt
@require_http_methods(["POST", "OPTIONS"])
def process_table(request):
    """
    POST /api/process/
    Body: JSON с полями sheet, range, hasHeaders, data
    Response: JSON с полями data, info
    """
    # CORS preflight
    if request.method == "OPTIONS":
        response = JsonResponse({})
        response["Access-Control-Allow-Origin"]  = "*"
        response["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        response["Access-Control-Allow-Headers"] = "Content-Type"
        return response

    try:
        body = json.loads(request.body)
    except json.JSONDecodeError as e:
        return _json_error(f"Неверный JSON: {e}", 400)

    data        = body.get("data")
    has_headers = body.get("hasHeaders", True)
    sheet       = body.get("sheet", "")
    range_str   = body.get("range", "")

    if not isinstance(data, list) or not data:
        return _json_error("Поле data должно быть непустым 2D-массивом", 400)

    try:
        result = run_macro(data, has_headers, sheet, range_str)
    except Exception as e:
        return _json_error(f"Ошибка макроса: {e}", 500)

    response = JsonResponse(result)
    response["Access-Control-Allow-Origin"] = "*"
    return response


@require_http_methods(["GET"])
def ping(request):
    """GET /api/ping/ — проверка доступности сервера"""
    response = JsonResponse({"status": "ok", "service": "Django Processor"})
    response["Access-Control-Allow-Origin"] = "*"
    return response


# ─────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────

def _json_error(message: str, status: int = 400) -> JsonResponse:
    r = JsonResponse({"error": message}, status=status)
    r["Access-Control-Allow-Origin"] = "*"
    return r
