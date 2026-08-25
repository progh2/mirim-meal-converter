(function () {
  "use strict";

  var dropzone = document.getElementById("dropzone");
  var fileInput = document.getElementById("file");
  var pickBtn = document.getElementById("pick-btn");
  var fileNameEl = document.getElementById("file-name");
  var errorEl = document.getElementById("error");
  var resultEl = document.getElementById("result");
  var summaryEl = document.getElementById("summary");
  var tbody = document.getElementById("preview-body");
  var downloadBtn = document.getElementById("download");

  var current = null;

  function showError(msg) {
    errorEl.hidden = false;
    errorEl.textContent = msg;
    resultEl.hidden = true;
    current = null;
  }

  function clearError() {
    errorEl.hidden = true;
    errorEl.textContent = "";
  }

  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function (ev) { resolve(ev.target.result); };
      reader.onerror = function () { reject(new Error("파일을 읽지 못했습니다.")); };
      reader.readAsArrayBuffer(file);
    });
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function render(result, originalName) {
    current = result;
    fileNameEl.hidden = false;
    fileNameEl.textContent = "올린 파일: " + originalName + "  →  " + result.filename;
    summaryEl.textContent =
      result.year + "년 " + result.month + "월 · " +
      result.records.length + "개 끼니 · 시트 " + (result.sheetName || "");

    tbody.textContent = "";
    result.records.forEach(function (rec) {
      var tr = document.createElement("tr");
      tr.innerHTML =
        "<td>" + escapeHtml(rec.dateStr) + "</td>" +
        "<td>" + escapeHtml(rec.mealName) + "</td>" +
        "<td>" + escapeHtml(rec.kcal) + "</td>" +
        "<td>" + escapeHtml(rec.firstDish || "") + "</td>";
      tbody.appendChild(tr);
    });
    resultEl.hidden = false;
  }

  async function handleFile(file) {
    clearError();
    resultEl.hidden = true;
    current = null;
    if (!file) return;
    var name = file.name || "식단표.xlsx";
    if (!/\.xlsx?$/i.test(name)) {
      showError("엑셀 파일(.xlsx 또는 .xls)만 올릴 수 있습니다.");
      return;
    }
    if (typeof XLSX === "undefined" || typeof MealConverter === "undefined") {
      showError("변환 스크립트를 불러오지 못했습니다. 페이지를 새로고침해 주세요.");
      return;
    }
    try {
      var buf = await readFile(file);
      var wb = XLSX.read(buf, { type: "array", cellDates: true });
      var result = MealConverter.parseWorkbook(XLSX, wb);
      render(result, name);
    } catch (err) {
      var msg = (err && err.message) ? err.message : String(err);
      if (err && err.name !== "MealPlanError") {
        msg = "식단표를 읽지 못했습니다. 월별 달력형 급식 식단표인지 확인해 주세요.";
      }
      showError(msg);
    }
  }

  function openPicker() {
    fileInput.click();
  }

  dropzone.addEventListener("click", function (e) {
    if (e.target === pickBtn) return;
    openPicker();
  });
  pickBtn.addEventListener("click", function (e) {
    e.stopPropagation();
    openPicker();
  });
  dropzone.addEventListener("keydown", function (e) {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      openPicker();
    }
  });
  fileInput.addEventListener("change", function () {
    if (fileInput.files && fileInput.files[0]) handleFile(fileInput.files[0]);
  });

  ["dragenter", "dragover"].forEach(function (type) {
    dropzone.addEventListener(type, function (e) {
      e.preventDefault();
      dropzone.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach(function (type) {
    dropzone.addEventListener(type, function (e) {
      e.preventDefault();
      dropzone.classList.remove("dragover");
    });
  });
  dropzone.addEventListener("drop", function (e) {
    var files = e.dataTransfer && e.dataTransfer.files;
    if (files && files[0]) handleFile(files[0]);
  });

  downloadBtn.addEventListener("click", async function () {
    if (!current) return;
    try {
      var out;
      if (typeof ExcelJS !== "undefined" && MealConverter.writeExcelJS) {
        out = await MealConverter.writeExcelJS(ExcelJS, current.records);
      } else {
        out = MealConverter.writeSheetJS(XLSX, current.records);
      }
      var blob = new Blob([out], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });
      var a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = current.filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
    } catch (err) {
      showError("엑셀 파일을 만들지 못했습니다. 다시 시도해 주세요.");
    }
  });
})();
