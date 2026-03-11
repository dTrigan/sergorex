"""
Django API для обработки данных из плагина Р7 Офис.

Добавь в urls.py:
    from .views import process_table_api
    urlpatterns = [
        ...
        path('api/process/', process_table_api, name='process_table'),
    ]

Добавь в settings.py (CORS для запросов из плагина):
    pip install django-cors-headers
    INSTALLED_APPS += ['corsheaders']
    MIDDLEWARE.insert(0, 'corsheaders.middleware.CorsMiddleware')
    CORS_ALLOW_ALL_ORIGINS = True  # или CORS_ALLOWED_ORIGINS = ['...']
"""

import json
import logging
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_http_methods

logger = logging.getLogger(__name__)


# ─── Реестр макросов ──────────────────────────────────────────────────────────

def macro_default(data: list[list], params: dict) -> dict:
    """
    Пример макроса: умножает все числовые ячейки на 2.
    Замени логику на свою.
    """
    result = []
    processed = 0

    for row in data:
        new_row = []
        for cell in row:
            try:
                val = float(cell)
                new_row.append(str(val * 2))
                processed += 1
            except (ValueError, TypeError):
                new_row.append(cell)
        result.append(new_row)

    return {
        "data": result,
        "stats": {"processed_rows": len(result), "processed_cells": processed},
    }


def macro_normalize(data: list[list], params: dict) -> dict:
    """
    Нормализует числовые столбцы в диапазон [0, 1].
    params: {"columns": [0, 2, 4]}  — индексы столбцов для нормализации
    """
    target_cols = params.get("columns", None)
    result = [row[:] for row in data]  # копия

    # Пропускаем строку заголовков если есть
    header_row = 1 if params.get("has_header", True) else 0
    data_rows = result[header_row:]

    col_count = max(len(r) for r in data_rows) if data_rows else 0

    for col_idx in range(col_count):
        if target_cols and col_idx not in target_cols:
            continue

        values = []
        for row in data_rows:
            if col_idx < len(row):
                try:
                    values.append(float(row[col_idx]))
                except (ValueError, TypeError):
                    pass

        if not values:
            continue

        col_min, col_max = min(values), max(values)
        spread = col_max - col_min or 1.0

        for row in data_rows:
            if col_idx < len(row):
                try:
                    val = float(row[col_idx])
                    row[col_idx] = f"{(val - col_min) / spread:.4f}"
                except (ValueError, TypeError):
                    pass

    return {
        "data": result,
        "stats": {"processed_rows": len(data_rows), "processed_cells": len(data_rows) * col_count},
    }


def macro_aggregate(data: list[list], params: dict) -> dict:
    """
    Агрегирует строки: возвращает сумму, среднее, мин, макс для числовых столбцов.
    """
    if not data:
        return {"data": [], "stats": {}}

    header = data[0] if params.get("has_header", True) else None
    rows = data[1:] if header else data

    col_count = max(len(r) for r in rows) if rows else 0
    sums = [0.0] * col_count
    counts = [0] * col_count
    mins = [float('inf')] * col_count
    maxs = [float('-inf')] * col_count

    for row in rows:
        for c in range(col_count):
            if c < len(row):
                try:
                    val = float(row[c])
                    sums[c] += val
                    counts[c] += 1
                    mins[c] = min(mins[c], val)
                    maxs[c] = max(maxs[c], val)
                except (ValueError, TypeError):
                    pass

    result = []
    if header:
        result.append(header)
    result.append(["Сумма"]    + [f"{sums[c]:.2f}"  if counts[c] else "" for c in range(col_count)])
    result.append(["Среднее"]  + [f"{sums[c]/counts[c]:.2f}" if counts[c] else "" for c in range(col_count)])
    result.append(["Минимум"]  + [f"{mins[c]:.2f}"  if counts[c] else "" for c in range(col_count)])
    result.append(["Максимум"] + [f"{maxs[c]:.2f}"  if counts[c] else "" for c in range(col_count)])

    return {"data": result, "stats": {"processed_rows": len(rows)}}


MACROS = {
    "default":   macro_default,
    "normalize": macro_normalize,
    "aggregate": macro_aggregate,
}


# ─── API endpoint ─────────────────────────────────────────────────────────────

@csrf_exempt
@require_http_methods(["POST", "OPTIONS"])
def process_table_api(request):
    """
    POST /api/process/

    Тело запроса (JSON):
        {
            "data":       [["A1","B1"], ["A2","B2"], ...],
            "rows":       2,
            "cols":       2,
            "macro":      "default",
            "params":     {},
            "sheet_name": "Лист1"
        }

    Ответ (JSON):
        {
            "success":   true,
            "data":      [["processed","cells"], ...],
            "start_row": 0,
            "start_col": 0,
            "stats":     {"processed_rows": N, ...}
        }
    """
    # OPTIONS — preflight CORS (если не используешь django-cors-headers)
    if request.method == "OPTIONS":
        response = JsonResponse({})
        response["Access-Control-Allow-Origin"] = "*"
        response["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        response["Access-Control-Allow-Headers"] = "Content-Type"
        return response

    try:
        body = json.loads(request.body)
    except json.JSONDecodeError as e:
        return _error(f"Невалидный JSON в теле запроса: {e}", 400)

    data = body.get("data")
    if not isinstance(data, list):
        return _error("Поле 'data' должно быть массивом", 400)

    macro_name = body.get("macro", "default")
    macro_fn = MACROS.get(macro_name)
    if macro_fn is None:
        return _error(f"Неизвестный макрос: '{macro_name}'. Доступны: {list(MACROS)}", 400)

    params = body.get("params", {})
    sheet_name = body.get("sheet_name", "")

    logger.info(
        "process_table_api: macro=%s sheet=%s rows=%d",
        macro_name, sheet_name, len(data)
    )

    try:
        result = macro_fn(data, params)
    except Exception as e:
        logger.exception("Ошибка выполнения макроса %s", macro_name)
        return _error(f"Ошибка выполнения макроса: {e}", 500)

    return JsonResponse({
        "success":   True,
        "data":      result["data"],
        "start_row": result.get("start_row", 0),
        "start_col": result.get("start_col", 0),
        "stats":     result.get("stats", {}),
    })


def _error(message: str, status: int = 400) -> JsonResponse:
    return JsonResponse({"success": False, "error": message}, status=status)
