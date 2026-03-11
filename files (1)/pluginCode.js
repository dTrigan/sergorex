"use strict";

// ─── Конфигурация ───────────────────────────────────────────────────────────
const PLUGIN_CONFIG = {
  SERVER_URL: "https://your-server.com/api/process/",  // ← замени на свой URL
  TIMEOUT_MS: 30000,
};

// ─── Состояние плагина ──────────────────────────────────────────────────────
let currentSheet = null;
let selectedRange = null;

// ─── Инициализация (вызывается Р7 при старте плагина) ───────────────────────
window.Asc.plugin.init = function () {
  // Получаем информацию о текущем листе и выделенном диапазоне
  window.Asc.plugin.callCommand(function () {
    const oWorkbook = Api.GetActiveSheet();
    const oRange = oWorkbook.GetSelection();
    const sAddress = oRange.GetAddress();
    return { sheetName: oWorkbook.GetName(), range: sAddress };
  }, false, function (result) {
    if (result) {
      currentSheet = result.sheetName;
      selectedRange = result.range;
      document.getElementById("info-range").textContent = result.range || "весь лист";
      document.getElementById("info-sheet").textContent = result.sheetName;
    }
  });
};

// ─── Читаем данные из таблицы ────────────────────────────────────────────────
function readTableData(useSelection) {
  return new Promise((resolve, reject) => {
    window.Asc.plugin.callCommand(function () {
      const sheet = Api.GetActiveSheet();
      let range;

      if (useSelection) {
        range = sheet.GetSelection();
      } else {
        // Определяем использованный диапазон (все данные листа)
        const usedRange = sheet.GetUsedRange();
        range = usedRange;
      }

      const rowCount = range.GetRowsCount();
      const colCount = range.GetColumnsCount();

      const data = [];
      for (let r = 0; r < rowCount; r++) {
        const row = [];
        for (let c = 0; c < colCount; c++) {
          const cell = range.GetCell(r, c);
          row.push(cell.GetText());
        }
        data.push(row);
      }

      return { data, rows: rowCount, cols: colCount };
    }, false, function (result) {
      if (result && result.data) {
        resolve(result);
      } else {
        reject(new Error("Не удалось прочитать данные из таблицы"));
      }
    });
  });
}

// ─── Записываем данные обратно в таблицу ────────────────────────────────────
function writeTableData(responseData, startRow, startCol) {
  return new Promise((resolve, reject) => {
    // Передаём данные через глобальную переменную (ограничение API Р7)
    window._pluginWriteData = {
      data: responseData.data,
      startRow: startRow || 0,
      startCol: startCol || 0,
    };

    window.Asc.plugin.callCommand(function () {
      const writeInfo = window._pluginWriteData;
      const sheet = Api.GetActiveSheet();

      for (let r = 0; r < writeInfo.data.length; r++) {
        for (let c = 0; c < writeInfo.data[r].length; c++) {
          const cell = sheet.GetRangeByNumber(
            writeInfo.startRow + r,
            writeInfo.startCol + c
          );
          const value = writeInfo.data[r][c];
          // Пытаемся записать число, иначе — строку
          const num = parseFloat(value);
          if (!isNaN(num) && String(num) === String(value)) {
            cell.SetValue(num);
          } else {
            cell.SetValue(value);
          }
        }
      }
    }, false, function () {
      delete window._pluginWriteData;
      resolve();
    });
  });
}

// ─── Отправка данных на Django-сервер ────────────────────────────────────────
async function sendToServer(tableData, macroName, extraParams) {
  const payload = {
    data: tableData.data,
    rows: tableData.rows,
    cols: tableData.cols,
    macro: macroName,
    params: extraParams || {},
    sheet_name: currentSheet,
  };

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), PLUGIN_CONFIG.TIMEOUT_MS);

  try {
    const response = await fetch(PLUGIN_CONFIG.SERVER_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Plugin-Version": "1.0.0",
      },
      body: JSON.stringify(payload),
      signal: controller.signal,
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Сервер вернул ${response.status}: ${errText}`);
    }

    return await response.json();
  } catch (err) {
    clearTimeout(timeoutId);
    if (err.name === "AbortError") {
      throw new Error("Превышено время ожидания ответа от сервера");
    }
    throw err;
  }
}

// ─── UI: обновление статуса ──────────────────────────────────────────────────
function setStatus(message, type) {
  // type: 'idle' | 'loading' | 'success' | 'error'
  const statusEl = document.getElementById("status-message");
  const statusIcon = document.getElementById("status-icon");
  const progressEl = document.getElementById("progress-bar");

  statusEl.textContent = message;
  statusEl.className = "status-text status-" + type;

  const icons = { idle: "○", loading: "◌", success: "✓", error: "✕" };
  statusIcon.textContent = icons[type] || "○";
  statusIcon.className = "status-icon icon-" + type;

  progressEl.style.display = type === "loading" ? "block" : "none";
}

function setButtonsDisabled(disabled) {
  document.getElementById("btn-process-all").disabled = disabled;
  document.getElementById("btn-process-selection").disabled = disabled;
}

// ─── Основной обработчик: запустить макрос ───────────────────────────────────
async function runProcess(useSelection) {
  const macroName = document.getElementById("macro-select").value;
  const writeMode = document.getElementById("write-mode").value;
  const extraParamsRaw = document.getElementById("extra-params").value.trim();

  let extraParams = {};
  if (extraParamsRaw) {
    try {
      extraParams = JSON.parse(extraParamsRaw);
    } catch {
      setStatus("Ошибка: дополнительные параметры — невалидный JSON", "error");
      return;
    }
  }

  setButtonsDisabled(true);
  setStatus("Читаю данные из таблицы...", "loading");

  try {
    // 1. Читаем данные из Р7
    const tableData = await readTableData(useSelection);
    setStatus(`Отправляю ${tableData.rows}×${tableData.cols} ячеек на сервер...`, "loading");

    // 2. Отправляем на Django
    const serverResponse = await sendToServer(tableData, macroName, extraParams);

    // 3. Проверяем ответ
    if (!serverResponse.success) {
      throw new Error(serverResponse.error || "Сервер сообщил об ошибке");
    }

    setStatus("Записываю результат в таблицу...", "loading");

    // 4. Записываем обратно
    const startRow = writeMode === "overwrite" ? 0 : (serverResponse.start_row || 0);
    const startCol = writeMode === "overwrite" ? 0 : (serverResponse.start_col || 0);

    await writeTableData(serverResponse, startRow, startCol);

    const stats = serverResponse.stats
      ? ` (обработано: ${serverResponse.stats.processed_rows} строк)`
      : "";
    setStatus(`Готово!${stats}`, "success");

  } catch (err) {
    setStatus("Ошибка: " + err.message, "error");
    console.error("[Django Plugin]", err);
  } finally {
    setButtonsDisabled(false);
  }
}

// ─── Привязка кнопок ─────────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", function () {
  document.getElementById("btn-process-all").addEventListener("click", () => {
    runProcess(false);
  });

  document.getElementById("btn-process-selection").addEventListener("click", () => {
    runProcess(true);
  });

  document.getElementById("btn-settings").addEventListener("click", () => {
    const panel = document.getElementById("settings-panel");
    panel.style.display = panel.style.display === "none" ? "block" : "none";
  });

  document.getElementById("server-url-input").addEventListener("change", function () {
    PLUGIN_CONFIG.SERVER_URL = this.value;
  });

  // Заполняем поле URL из конфига
  document.getElementById("server-url-input").value = PLUGIN_CONFIG.SERVER_URL;
});
