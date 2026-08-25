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
  var zeroKcalEl = document.getElementById("zero-kcal");
  var schoolForm = document.getElementById("neis-form");
  var schoolQ = document.getElementById("school-q");
  var schoolResults = document.getElementById("school-results");
  var schoolPicked = document.getElementById("school-picked");
  var monthInput = document.getElementById("neis-month");
  var fetchBtn = document.getElementById("neis-fetch");

  var current = null;
  var selectedSchool = null;

  function defaultSchool() {
    var d = (typeof MealConverter !== "undefined" && MealConverter.DEFAULT_NEIS_SCHOOL) || {
      SCHUL_NM: "미림마이스터고등학교",
      ATPT_OFCDC_SC_CODE: "B10",
      SD_SCHUL_CODE: "7011569",
      ATPT_OFCDC_SC_NM: "서울특별시교육청"
    };
    return {
      SCHUL_NM: d.SCHUL_NM,
      ATPT_OFCDC_SC_CODE: d.ATPT_OFCDC_SC_CODE,
      SD_SCHUL_CODE: d.SD_SCHUL_CODE,
      ATPT_OFCDC_SC_NM: d.ATPT_OFCDC_SC_NM
    };
  }

  function pad2(n) {
    return (n < 10 ? "0" : "") + n;
  }

  function showError(msg) {
    errorEl.hidden = false;
    errorEl.textContent = msg;
    resultEl.hidden = true;
    fileNameEl.hidden = true;
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

  function schoolLabel(school) {
    return school.SCHUL_NM + " · " + (school.ATPT_OFCDC_SC_NM || "");
  }

  function renderPicked() {
    if (!selectedSchool) {
      schoolPicked.textContent = "목록에서 학교를 선택해 주세요.";
      schoolPicked.classList.remove("is-set");
      return;
    }
    schoolPicked.textContent = "선택한 학교: " + schoolLabel(selectedSchool);
    schoolPicked.classList.add("is-set");
  }

  function renderSchoolResults(result) {
    schoolResults.textContent = "";
    if (!result || !result.schools || !result.schools.length) {
      schoolResults.hidden = true;
      return;
    }
    if (result.truncated) {
      var note = document.createElement("p");
      note.className = "school-note";
      note.textContent = "검색 결과가 많습니다. 학교 이름을 더 구체적으로 입력해 주세요.";
      schoolResults.appendChild(note);
    }
    var list = document.createElement("ul");
    list.className = "school-list";
    result.schools.forEach(function (school) {
      var li = document.createElement("li");
      var btn = document.createElement("button");
      var selected = selectedSchool &&
        selectedSchool.ATPT_OFCDC_SC_CODE === school.ATPT_OFCDC_SC_CODE &&
        selectedSchool.SD_SCHUL_CODE === school.SD_SCHUL_CODE;
      btn.type = "button";
      btn.className = "school-item" + (selected ? " is-selected" : "");
      btn.textContent = schoolLabel(school);
      btn.addEventListener("click", function () {
        selectedSchool = school;
        renderPicked();
        renderSchoolResults(result);
      });
      li.appendChild(btn);
      list.appendChild(li);
    });
    schoolResults.appendChild(list);
    schoolResults.hidden = false;
  }

  function render(result, sourceLabel) {
    current = result;
    fileNameEl.hidden = false;
    fileNameEl.textContent = sourceLabel + "  →  " + result.filename;
    summaryEl.textContent =
      result.year + "년 " + result.month + "월 · " +
      result.records.length + "개 끼니" +
      (result.sheetName ? " · " + result.sheetName : "");

    fillPreview();
    resultEl.hidden = false;
  }

  function writeOptions() {
    return { zeroKcal: !!(zeroKcalEl && zeroKcalEl.checked) };
  }

  function recordsForDownload(records, options) {
    if (!(options && options.zeroKcal)) return records;
    return records.map(function (rec) {
      var copy = {};
      for (var key in rec) {
        if (Object.prototype.hasOwnProperty.call(rec, key)) copy[key] = rec[key];
      }
      copy.kcal = 0;
      return copy;
    });
  }

  function fillPreview() {
    if (!current) return;
    var zero = writeOptions().zeroKcal;
    tbody.textContent = "";
    current.records.forEach(function (rec) {
      var tr = document.createElement("tr");
      tr.innerHTML =
        "<td>" + escapeHtml(rec.dateStr) + "</td>" +
        "<td>" + escapeHtml(rec.mealName) + "</td>" +
        "<td>" + escapeHtml(zero ? 0 : rec.kcal) + "</td>" +
        "<td class=\"menu-cell\">" + escapeHtml(rec.menu || "") + "</td>";
      tbody.appendChild(tr);
    });
  }

  function koreanError(err, fallback) {
    var msg = (err && err.message) ? err.message : String(err);
    if (err && err.name === "MealPlanError") return msg;
    return fallback;
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
      render(result, "올린 파일: " + name);
    } catch (err) {
      showError(koreanError(err, "식단표를 읽지 못했습니다. 월별 달력형 급식 식단표인지 확인해 주세요."));
    }
  }

  function openPicker() {
    fileInput.click();
  }

  async function searchSchools(ev) {
    if (ev) ev.preventDefault();
    clearError();
    if (typeof MealConverter === "undefined") {
      showError("변환 스크립트를 불러오지 못했습니다. 페이지를 새로고침해 주세요.");
      return;
    }
    var q = schoolQ.value.trim();
    try {
      var result = await MealConverter.searchNeisSchools(q);
      if (result.schools.length === 1) {
        selectedSchool = result.schools[0];
      } else if (selectedSchool) {
        var keep = result.schools.some(function (s) {
          return s.ATPT_OFCDC_SC_CODE === selectedSchool.ATPT_OFCDC_SC_CODE &&
            s.SD_SCHUL_CODE === selectedSchool.SD_SCHUL_CODE;
        });
        if (!keep) selectedSchool = null;
      }
      renderPicked();
      renderSchoolResults(result);
    } catch (err) {
      schoolResults.hidden = true;
      selectedSchool = null;
      renderPicked();
      showError(koreanError(err, "학교를 검색하지 못했습니다. 잠시 후 다시 시도해 주세요."));
    }
  }

  async function fetchMeals() {
    clearError();
    resultEl.hidden = true;
    current = null;
    if (typeof MealConverter === "undefined") {
      showError("변환 스크립트를 불러오지 못했습니다. 페이지를 새로고침해 주세요.");
      return;
    }
    if (!selectedSchool) {
      showError("학교를 선택한 뒤 급식을 가져와 주세요.");
      return;
    }
    var ym = (monthInput.value || "").split("-");
    var year = +ym[0];
    var month = +ym[1];
    if (!year || !month) {
      showError("연월을 선택해 주세요.");
      return;
    }
    fetchBtn.disabled = true;
    fetchBtn.textContent = "가져오는 중…";
    try {
      var result = await MealConverter.fetchNeisMeals({
        ATPT_OFCDC_SC_CODE: selectedSchool.ATPT_OFCDC_SC_CODE,
        SD_SCHUL_CODE: selectedSchool.SD_SCHUL_CODE,
        year: year,
        month: month
      });
      render(result, "나이스: " + selectedSchool.SCHUL_NM);
    } catch (err) {
      showError(koreanError(err, "나이스 급식 정보를 가져오지 못했습니다. 잠시 후 다시 시도해 주세요."));
    } finally {
      fetchBtn.disabled = false;
      fetchBtn.textContent = "급식 가져오기";
    }
  }

  selectedSchool = defaultSchool();
  renderPicked();
  if (monthInput && !monthInput.value) {
    var now = new Date();
    monthInput.value = now.getFullYear() + "-" + pad2(now.getMonth() + 1);
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

  schoolForm.addEventListener("submit", searchSchools);
  fetchBtn.addEventListener("click", fetchMeals);

  if (zeroKcalEl) {
    zeroKcalEl.addEventListener("change", function () {
      if (current) fillPreview();
    });
  }

  downloadBtn.addEventListener("click", async function () {
    if (!current) return;
    try {
      var out;
      var options = writeOptions();
      var records = recordsForDownload(current.records, options);
      if (typeof ExcelJS !== "undefined" && MealConverter.writeExcelJS) {
        out = await MealConverter.writeExcelJS(ExcelJS, records, options);
      } else {
        out = MealConverter.writeSheetJS(XLSX, records, options);
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
