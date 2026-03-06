# sergorex
LUCKY<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8" />
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: sans-serif;
      background: #1a1a2e;
      color: #e0e0e0;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      height: 100vh;
      gap: 16px;
      padding: 20px;
    }
    #sendBtn {
      width: 100%;
      padding: 14px;
      background: #3b82f6;
      border: none;
      border-radius: 10px;
      color: #fff;
      font-size: 15px;
      font-weight: 700;
      cursor: pointer;
      transition: background 0.2s, transform 0.1s;
    }
    #sendBtn:hover:not(:disabled) { background: #2563eb; }
    #sendBtn:active:not(:disabled) { transform: scale(0.97); }
    #sendBtn:disabled { opacity: 0.5; cursor: not-allowed; }
    #status {
      font-size: 12px;
      text-align: center;
      min-height: 18px;
      color: #94a3b8;
    }
    #status.ok  { color: #22c55e; }
    #status.err { color: #ef4444; }
  </style>
</head>
<body>

  <button id="sendBtn" onclick="sendFile()">📤 Отправить файл на сервер</button>
  <div id="status">Нажмите кнопку для отправки</div>

  <script src="https://onlyoffice.github.io/sdkjs-plugins/v1/plugins.js"></script>
  <script src="https://onlyoffice.github.io/sdkjs-plugins/v1/plugins-ui.js"></script>
  <script>
    // URL сервера — замените на свой
    var SERVER_URL = "https://httpbin.org/post";

    window.Asc.plugin.init = function () {};
    window.Asc.plugin.button = function () {};

    function sendFile() {
      var btn = document.getElementById("sendBtn");
      var status = document.getElementById("status");

      btn.disabled = true;
      status.className = "";
      status.textContent = "Получение файла...";

      // Запрашиваем текущий файл в base64 через API плагина
      window.Asc.plugin.executeMethod("GetFileData", [], function (fileData) {
        if (!fileData) {
          status.textContent = "Ошибка: не удалось получить файл";
          status.className = "err";
          btn.disabled = false;
          return;
        }

        status.textContent = "Отправка на сервер...";

        // Конвертируем base64 → Blob
        var binary = atob(fileData);
        var bytes = new Uint8Array(binary.length);
        for (var i = 0; i < binary.length; i++) {
          bytes[i] = binary.charCodeAt(i);
        }
        var blob = new Blob([bytes], { type: "application/octet-stream" });

        // Формируем multipart/form-data
        var formData = new FormData();
        formData.append("file", blob, "document.xlsx");

        fetch(SERVER_URL, {
          method: "POST",
          body: formData
        })
        .then(function (res) {
          if (res.ok) {
            status.textContent = "✓ Файл успешно отправлен! (HTTP " + res.status + ")";
            status.className = "ok";
          } else {
            status.textContent = "✗ Ошибка сервера: " + res.status;
            status.className = "err";
          }
        })
        .catch(function (err) {
          status.textContent = "✗ " + err.message;
          status.className = "err";
        })
        .finally(function () {
          btn.disabled = false;
        });
      });
    }
  </script>
</body>
</html>
