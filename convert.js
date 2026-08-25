/**
 * 월별 식단표(엑셀 또는 나이스) → 연구정보원 학교 홈페이지 mlsvTmplat 변환
 * Browser + Node (no DOM).
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === "object" && module.exports) {
    module.exports = api;
  }
  if (typeof window !== "undefined") {
    window.MealConverter = api;
  }
  root.MealConverter = api;
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";
  var MEAL_NAMES = ["조식", "중식", "석식"];
  var MEAL_CODE = { 조식: 0, 중식: 1, 석식: 2 };
  var NUTR_RE = /^\s*\d+(\.\d+)?\s*\/\s*\d/;
  var DATE_TEXT_RE = /(\d{1,2})\s*월\s*(\d{1,2})\s*일/;
  var DATE_ISO_RE = /(\d{4})\s*[.\-\/년]\s*(\d{1,2})\s*[.\-\/월]\s*(\d{1,2})/;
  var EVENT_RE = /이벤트|의\s*날|급식의\s*날|응모|발표/;
  var SKIP_EXACT = {
    영양표시제: 1,
    "(주간평균)": 1,
    주간평균: 1,
    "열량/단백질/칼슘/철분": 1,
    "열량/단백질/칼슔/철분": 1
  };

  function MealPlanError(message) {
    this.name = "MealPlanError";
    this.message = message;
    if (typeof Error.captureStackTrace === "function") {
      Error.captureStackTrace(this, MealPlanError);
    }
  }
  MealPlanError.prototype = Object.create(Error.prototype);
  MealPlanError.prototype.constructor = MealPlanError;

  function isDateObj(v) {
    return Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime());
  }

  function cellStr(v) {
    if (v == null) return "";
    if (isDateObj(v)) return "";
    return String(v).replace(/\u00a0/g, " ").replace(/\r\n/g, "\n").trim();
  }

  function isBlank(s) {
    return !s || /^[\s\-\u3000.]*$/.test(s);
  }

  function isSkipText(s) {
    if (isBlank(s)) return true;
    if (SKIP_EXACT[s]) return true;
    if (/영양표시제/.test(s)) return true;
    if (/주간평균/.test(s)) return true;
    if (/원산지\s*표시/.test(s) || /원산지\s*안내/.test(s)) return true;
    if (/식품알레르기|알레르기\s*표기/.test(s)) return true;
    if (/^열량\s*:/.test(s) || /^단백질\s*:/.test(s)) return true;
    if (/^칼슘\s*:/.test(s) || /^칼슔\s*:/.test(s) || /^철분\s*:/.test(s)) return true;
    if (/^열량\s*\/\s*단백질/.test(s)) return true;
    return false;
  }

  function isNutrition(v) {
    if (v == null || isDateObj(v)) return false;
    var t = cellStr(v);
    if (!t) return false;
    t = t.replace(/\s+/g, "");
    return NUTR_RE.test(t);
  }

  function neatRound(n) {
    var r = Math.round(n * 10) / 10;
    return Math.abs(r - Math.round(r)) < 1e-8 ? Math.round(r) : r;
  }

  // Excel Kcal is 열량 only (first number). Do not add 단백질/칼슘/철분.
  function parseKcal(v) {
    if (!isNutrition(v)) return null;
    var t = cellStr(v).replace(/\s+/g, "");
    var parts = t.split("/").filter(Boolean);
    var i, n;
    for (i = 0; i < parts.length; i++) {
      n = parseFloat(parts[i]);
      if (isFinite(n)) return neatRound(n);
    }
    return null;
  }

  function isPureEventPart(part) {
    part = String(part || "").trim();
    if (!part) return true;
    if (!EVENT_RE.test(part)) return false;
    if (/\d+\.\d+/.test(part)) return false;
    return true;
  }

  function isEventText(s) {
    if (isBlank(s) || isSkipText(s) || isNutrition(s)) return false;
    if (!EVENT_RE.test(s)) return false;
    return String(s).split("/").every(isPureEventPart);
  }

  function isDishText(s) {
    if (isBlank(s) || isSkipText(s) || isNutrition(s)) return false;
    if (isEventText(s)) return false;
    return true;
  }

  function toExcelSerial(y, m, d) {
    var utc = Date.UTC(y, m - 1, d);
    var epoch = Date.UTC(1899, 11, 30);
    return Math.round((utc - epoch) / 86400000);
  }

  function excelSerialToYMD(n) {
    var serial = Math.floor(Number(n));
    var utc = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
    return {
      y: utc.getUTCFullYear(),
      m: utc.getUTCMonth() + 1,
      d: utc.getUTCDate()
    };
  }

  function ymdFromDate(v) {
    var utcH = v.getUTCHours();
    var locH = v.getHours();
    if (utcH === 0) {
      return { y: v.getUTCFullYear(), m: v.getUTCMonth() + 1, d: v.getUTCDate() };
    }
    if (locH === 0) {
      return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
    }
    if (utcH >= 20) {
      var next = new Date(v.getTime() + (24 - utcH) * 3600 * 1000);
      return { y: next.getUTCFullYear(), m: next.getUTCMonth() + 1, d: next.getUTCDate() };
    }
    return { y: v.getUTCFullYear(), m: v.getUTCMonth() + 1, d: v.getUTCDate() };
  }

  function parseDate(v, fallbackYear) {
    if (v == null || v === "") return null;
    if (isDateObj(v)) return ymdFromDate(v);
    if (typeof v === "number" && isFinite(v) && v > 20000 && v < 120000) {
      return excelSerialToYMD(v);
    }
    var t = cellStr(v);
    if (!t) return null;
    var m = t.match(DATE_ISO_RE);
    if (m) return { y: +m[1], m: +m[2], d: +m[3] };
    m = t.match(DATE_TEXT_RE);
    if (m) {
      return {
        y: fallbackYear || new Date().getFullYear(),
        m: +m[1],
        d: +m[2]
      };
    }
    m = t.match(/^(\d{1,2})[.\-\/](\d{1,2})$/);
    if (m && fallbackYear) {
      return { y: fallbackYear, m: +m[1], d: +m[2] };
    }
    return null;
  }

  function pad2(n) {
    return (n < 10 ? "0" : "") + n;
  }

  function formatYMD(p) {
    return p.y + "-" + pad2(p.m) + "-" + pad2(p.d);
  }

  function weekday(p) {
    return new Date(p.y, p.m - 1, p.d).getDay();
  }

  function gridCell(grid, r, c) {
    var row = grid[r];
    if (!row) return null;
    return row[c];
  }

  function rowMealLabel(grid, r) {
    var row = grid[r] || [];
    var max = Math.min(row.length, 4);
    for (var c = 0; c < max; c++) {
      var s = cellStr(row[c]);
      if (MEAL_CODE.hasOwnProperty(s)) return s;
    }
    return null;
  }
  function inferContext(grid, sheetName) {
    var year = null;
    var month = null;
    var i, j, v, d, s, m;
    for (i = 0; i < grid.length; i++) {
      var row = grid[i] || [];
      for (j = 0; j < row.length; j++) {
        v = row[j];
        if (isDateObj(v) || (typeof v === "number" && v > 20000)) {
          d = parseDate(v);
          if (d) {
            if (!year) year = d.y;
            if (!month) month = d.m;
          }
        }
        s = cellStr(v);
        if (s) {
          m = s.match(/(\d{4})\s*년/);
          if (m && !year) year = +m[1];
          m = s.match(/(\d{1,2})\s*월\s*식단/);
          if (m && !month) month = +m[1];
        }
      }
    }
    s = String(sheetName || "");
    m = s.match(/(\d{4})/);
    if (m && !year) year = +m[1];
    m = s.match(/(\d{1,2})\s*월/);
    if (m && !month) month = +m[1];
    return {
      year: year || new Date().getFullYear(),
      month: month || null
    };
  }

  function findDateRows(grid, year) {
    var rows = [];
    for (var r = 0; r < grid.length; r++) {
      if (rowMealLabel(grid, r)) continue;
      var row = grid[r] || [];
      var dates = [];
      for (var c = 0; c < row.length; c++) {
        var d = parseDate(row[c], year);
        if (d && d.d >= 1 && d.d <= 31 && d.m >= 1 && d.m <= 12) {
          dates.push({ c: c, d: d });
        }
      }
      if (dates.length >= 2) {
        rows.push({ r: r, dates: dates });
      }
    }
    return rows;
  }

  function classifyRow(grid, r, dateCols) {
    var hasDish = false;
    var hasEvent = false;
    var hasNutr = false;
    var i, v, s;
    for (i = 0; i < dateCols.length; i++) {
      v = gridCell(grid, r, dateCols[i]);
      if (isNutrition(v)) {
        hasNutr = true;
        continue;
      }
      s = cellStr(v);
      if (isSkipText(s)) continue;
      if (isEventText(s)) {
        hasEvent = true;
        continue;
      }
      if (isDishText(s)) hasDish = true;
    }
    return {
      label: rowMealLabel(grid, r),
      hasDish: hasDish,
      hasEvent: hasEvent,
      hasNutr: hasNutr
    };
  }

  function isContinuationRow(grid, r, dateCols, weekStart, weekEnd) {
    if (r <= weekStart || r >= weekEnd) return false;
    if (rowMealLabel(grid, r)) return false;
    var info = classifyRow(grid, r, dateCols);
    if (info.hasNutr) return false;
    return info.hasDish || info.hasEvent;
  }

  function collectMeal(grid, rows, col) {
    var items = [];
    var events = [];
    var kcal = 0;
    var seen = {};
    var i, r, v, s, k, t, lines, li, line;
    function addItem(text) {
      t = text.replace(/[ \t]+/g, " ").trim();
      if (isBlank(t) || isSkipText(t) || isNutrition(t)) return;
      if (seen[t]) return;
      seen[t] = 1;
      if (isEventText(t)) events.push(t);
      else items.push(t);
    }
    for (i = 0; i < rows.length; i++) {
      r = rows[i];
      v = gridCell(grid, r, col);
      if (isNutrition(v)) {
        k = parseKcal(v);
        if (k != null) kcal = k;
        continue;
      }
      s = cellStr(v);
      if (isSkipText(s)) continue;
      lines = s.split("\n");
      for (li = 0; li < lines.length; li++) {
        line = lines[li].replace(/[ \t]+/g, " ").trim();
        if (line) addItem(line);
      }
    }
    for (i = 0; i < events.length; i++) {
      t = events[i];
      if (items.some(function (it) { return it.indexOf(t) !== -1; })) continue;
      items.push(t);
    }
    return { items: items, kcal: kcal };
  }

  function parseMealPlan(grid, options) {
    options = options || {};
    if (!grid || !grid.length) {
      throw new MealPlanError(
        "이 파일은 급식 식단표가 아닌 것 같습니다. 월별 달력형 식단표(조식·중식·석식) 엑셀을 올려 주세요."
      );
    }

    var ctx = inferContext(grid, options.sheetName);
    var dateRows = findDateRows(grid, ctx.year);
    if (!dateRows.length) {
      throw new MealPlanError(
        "날짜가 있는 식단표를 찾지 못했습니다. 달력형 월별 급식 식단표(.xlsx/.xls)인지 확인해 주세요."
      );
    }

    var records = [];
    var wi, week, nextWeek, weekStart, weekEnd, dateCols, colDates;
    var meals, r, label, mi, meal, start, end, look, rows;
    var di, col, date, collected, rec;

    for (wi = 0; wi < dateRows.length; wi++) {
      week = dateRows[wi];
      nextWeek = dateRows[wi + 1];
      weekStart = week.r;
      weekEnd = nextWeek ? nextWeek.r : grid.length;
      dateCols = [];
      colDates = {};
      for (di = 0; di < week.dates.length; di++) {
        col = week.dates[di].c;
        date = week.dates[di].d;
        date.y = date.y || ctx.year;
        if (weekday(date) === 0 || weekday(date) === 6) continue;
        dateCols.push(col);
        colDates[col] = date;
      }
      if (!dateCols.length) continue;

      meals = [];
      for (r = weekStart + 1; r < weekEnd; r++) {
        label = rowMealLabel(grid, r);
        if (label) meals.push({ name: label, labelRow: r });
      }
      if (!meals.length) continue;

      for (mi = 0; mi < meals.length; mi++) {
        meal = meals[mi];
        start = meal.labelRow;
        look = start - 1;
        if (isContinuationRow(grid, look, dateCols, weekStart, weekEnd)) {
          start = look;
        }
        meal.start = start;
      }
      for (mi = 0; mi < meals.length; mi++) {
        meal = meals[mi];
        end = mi + 1 < meals.length ? meals[mi + 1].start - 1 : weekEnd - 1;
        meal.end = end;
        rows = [];
        for (r = meal.start; r <= meal.end; r++) rows.push(r);

        for (di = 0; di < dateCols.length; di++) {
          col = dateCols[di];
          date = colDates[col];
          collected = collectMeal(grid, rows, col);
          if (!collected.items.length) continue;
          rec = {
            y: date.y,
            m: date.m,
            d: date.d,
            dateStr: formatYMD(date),
            mealCode: MEAL_CODE[meal.name],
            mealName: meal.name,
            menu: collected.items.join("\n"),
            kcal: collected.kcal == null ? 0 : collected.kcal,
            firstDish: collected.items[0]
          };
          records.push(rec);
        }
      }
    }

    if (!records.length) {
      throw new MealPlanError(
        "식단 메뉴를 읽지 못했습니다. 조식·중식·석식 칸이 있는 달력형 월 식단표인지 확인해 주세요."
      );
    }

    records.sort(function (a, b) {
      if (a.y !== b.y) return a.y - b.y;
      if (a.m !== b.m) return a.m - b.m;
      if (a.d !== b.d) return a.d - b.d;
      return a.mealCode - b.mealCode;
    });

    var year = records[0].y;
    var month = records[0].m;
    return {
      records: records,
      year: year,
      month: month,
      filename: "mlsvTmplat_" + year + pad2(month) + ".xlsx"
    };
  }
  var HEADERS = [
    "년월일\n(YYYYMMDD)",
    "아침:0 점심:1\n저녁:2(빈값:1)",
    "제목\n(빈값:B컬럼)",
    "식단\n(빈값:메뉴없음)",
    "Kcal\n(빈값:0)"
  ];

  function pickMealSheet(XLSX, workbook) {
    var names = workbook.SheetNames || [];
    var best = null;
    var bestScore = -1;
    var i, name, ws, grid, result, score, text;
    for (i = 0; i < names.length; i++) {
      name = names[i];
      ws = workbook.Sheets[name];
      if (!ws) continue;
      grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
      text = name + " " + grid.slice(0, 8).map(function (row) {
        return (row || []).join(" ");
      }).join(" ");
      score = 0;
      if (/\d+\s*월/.test(name) || /식단/.test(name)) score += 5;
      if (/조식/.test(text) && /중식/.test(text)) score += 8;
      if (/식단표/.test(text)) score += 4;
      if (/미림/.test(text)) score += 2;
      try {
        result = parseMealPlan(grid, { sheetName: name });
        score += Math.min(result.records.length, 80);
        if (score > bestScore) {
          bestScore = score;
          best = result;
          best.sheetName = name;
        }
      } catch (e) {
        /* try next sheet */
      }
    }
    if (!best) {
      throw new MealPlanError(
        "이 파일은 급식 식단표가 아닌 것 같습니다. 월별 달력형 식단표(조식·중식·석식) 엑셀을 올려 주세요."
      );
    }
    return best;
  }

  function parseWorkbook(XLSX, workbook) {
    return pickMealSheet(XLSX, workbook);
  }

  function kcalForWrite(rec, options) {
    if (options && options.zeroKcal) return 0;
    return rec && rec.kcal != null ? rec.kcal : 0;
  }

  function applySheetJSSheet(XLSX, records, options) {
    var aoa = [HEADERS];
    var i, rec;
    for (i = 0; i < records.length; i++) {
      rec = records[i];
      aoa.push([
        toExcelSerial(rec.y, rec.m, rec.d),
        rec.mealCode,
        rec.mealName,
        rec.menu,
        kcalForWrite(rec, options)
      ]);
    }
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [
      { wch: 14 },
      { wch: 15 },
      { wch: 10 },
      { wch: 62 },
      { wch: 10 }
    ];
    ws["!rows"] = [{ hpt: 32 }];
    var range = XLSX.utils.decode_range(ws["!ref"]);
    for (var r = range.s.r; r <= range.e.r; r++) {
      for (var c = range.s.c; c <= range.e.c; c++) {
        var addr = XLSX.utils.encode_cell({ r: r, c: c });
        var cell = ws[addr];
        if (!cell) continue;
        if (r === 0) {
          cell.s = cell.s || {};
          cell.s.alignment = { wrapText: true, horizontal: "center", vertical: "center" };
          cell.s.font = { name: "Arial", sz: 10, bold: true };
        } else if (c === 0) {
          cell.t = "n";
          cell.z = "yyyymmdd";
        } else if (c === 3) {
          cell.s = cell.s || {};
          cell.s.alignment = { wrapText: true, vertical: "center" };
          cell.s.font = { name: "Arial", sz: 10 };
        } else {
          cell.s = cell.s || {};
          cell.s.font = { name: "Arial", sz: 10 };
        }
      }
    }
    return ws;
  }

  function writeSheetJS(XLSX, records, options) {
    var wb = XLSX.utils.book_new();
    var ws = applySheetJSSheet(XLSX, records, options);
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    return XLSX.write(wb, { bookType: "xlsx", type: "array" });
  }

  var NEIS_HUB = "https://open.neis.go.kr/hub";
  var DEFAULT_NEIS_SCHOOL = {
    SCHUL_NM: "미림마이스터고등학교",
    searchName: "미림마이스터고",
    ATPT_OFCDC_SC_CODE: "B10",
    SD_SCHUL_CODE: "7011569",
    ATPT_OFCDC_SC_NM: "서울특별시교육청"
  };
  var NEIS_MEAL = {
    "1": { code: 0, name: "조식" },
    "2": { code: 1, name: "중식" },
    "3": { code: 2, name: "석식" }
  };

  function lastDayOfMonth(year, month) {
    return new Date(year, month, 0).getDate();
  }

  // NEIS Kcal is CAL_INFO only. Do not add NTR_INFO — 칼슘 is mg.
  function parseNeisKcal(calInfo) {
    if (calInfo == null || !String(calInfo).trim()) return 0;
    var cm = String(calInfo).replace(/,/g, "").match(/([0-9]+(?:\.[0-9]+)?)/);
    if (!cm) return 0;
    return neatRound(parseFloat(cm[1]));
  }

  function formatNeisDish(s) {
    s = String(s || "").replace(/\u00a0/g, " ").replace(/[ \t]+/g, " ").trim();
    s = s.replace(/\s*\(\s*j\s*\)/gi, "");
    s = s.replace(/\s*\(\s*(조식|중식|석식)\s*\)/g, "");
    s = s.replace(/[ \t]+/g, " ").trim();
    var m = s.match(/^(.*?)\s*\((\d+(?:\.\d+)*)\)\s*$/);
    if (m) return m[1].replace(/[ \t]+/g, " ").trim() + m[2];
    return s;
  }

  function dishesFromNeis(ddishNm) {
    if (ddishNm == null) return [];
    return String(ddishNm)
      .split(/<br\s*\/?>/i)
      .map(formatNeisDish)
      .filter(function (s) {
        return !isBlank(s);
      });
  }

  function parseYmdCompact(s) {
    var t = String(s || "").replace(/\D/g, "");
    if (t.length !== 8) return null;
    return { y: +t.slice(0, 4), m: +t.slice(4, 6), d: +t.slice(6, 8) };
  }

  function recordFromNeisRow(row) {
    if (!row) return null;
    var meal = NEIS_MEAL[String(row.MMEAL_SC_CODE)];
    if (!meal) return null;
    var date = parseYmdCompact(row.MLSV_YMD);
    if (!date) return null;
    var items = dishesFromNeis(row.DDISH_NM);
    if (!items.length) return null;
    return {
      y: date.y,
      m: date.m,
      d: date.d,
      dateStr: formatYMD(date),
      mealCode: meal.code,
      mealName: meal.name,
      menu: items.join("\n"),
      kcal: parseNeisKcal(row.CAL_INFO),
      firstDish: items[0]
    };
  }

  function recordsFromNeisRows(rows) {
    var records = [];
    var i, rec;
    rows = rows || [];
    for (i = 0; i < rows.length; i++) {
      rec = recordFromNeisRow(rows[i]);
      if (rec) records.push(rec);
    }
    records.sort(function (a, b) {
      if (a.y !== b.y) return a.y - b.y;
      if (a.m !== b.m) return a.m - b.m;
      if (a.d !== b.d) return a.d - b.d;
      return a.mealCode - b.mealCode;
    });
    if (!records.length) {
      throw new MealPlanError(
        "이 달의 급식 정보가 없습니다. 아직 나이스에 공개되지 않은 달일 수 있습니다."
      );
    }
    var year = records[0].y;
    var month = records[0].m;
    return {
      records: records,
      year: year,
      month: month,
      filename: "mlsvTmplat_" + year + pad2(month) + ".xlsx",
      sheetName: "나이스"
    };
  }

  function neisUrl(path, query) {
    var keys = Object.keys(query);
    var q = keys.map(function (k) {
      var v = query[k] == null ? "" : query[k];
      return encodeURIComponent(k) + "=" + encodeURIComponent(v);
    }).join("&");
    return NEIS_HUB + "/" + path + "?" + q;
  }

  function readNeisPayload(data, listKey) {
    if (!data || typeof data !== "object") {
      return { code: "ERROR", message: "나이스 응답을 읽지 못했습니다.", total: 0, rows: [] };
    }
    if (data.RESULT && data.RESULT.CODE) {
      return {
        code: data.RESULT.CODE,
        message: data.RESULT.MESSAGE || "",
        total: 0,
        rows: []
      };
    }
    var block = data[listKey];
    if (!Array.isArray(block) || !block.length) {
      return { code: "ERROR", message: "나이스 응답 형식이 올바르지 않습니다.", total: 0, rows: [] };
    }
    var head = (block[0] && block[0].head) || [];
    var code = "INFO-000";
    var message = "";
    var total = 0;
    var rows = [];
    var i, item;
    for (i = 0; i < head.length; i++) {
      item = head[i] || {};
      if (item.list_total_count != null) total = Number(item.list_total_count) || 0;
      if (item.RESULT && item.RESULT.CODE) {
        code = item.RESULT.CODE;
        message = item.RESULT.MESSAGE || "";
      }
    }
    for (i = 0; i < block.length; i++) {
      if (block[i] && Array.isArray(block[i].row)) {
        rows = rows.concat(block[i].row);
      }
    }
    return { code: code, message: message, total: total, rows: rows };
  }

  function getFetch(opts) {
    if (opts && opts.fetch) return opts.fetch;
    if (typeof fetch === "function") return fetch;
    throw new MealPlanError("이 브라우저에서는 나이스 조회를 사용할 수 없습니다.");
  }

  async function fetchNeisJson(url, fetchFn) {
    var res;
    try {
      res = await fetchFn(url);
    } catch (e) {
      throw new MealPlanError("나이스 교육정보 개방 포털에 연결하지 못했습니다. 잠시 후 다시 시도해 주세요.");
    }
    if (!res || !res.ok) {
      throw new MealPlanError("나이스 교육정보 개방 포털에 연결하지 못했습니다. 잠시 후 다시 시도해 주세요.");
    }
    try {
      return await res.json();
    } catch (e) {
      throw new MealPlanError("나이스 응답을 읽지 못했습니다.");
    }
  }

  function mealQuery(office, school, fromYmd, toYmd, page) {
    return neisUrl("mealServiceDietInfo", {
      Type: "json",
      pIndex: page || 1,
      pSize: 100,
      ATPT_OFCDC_SC_CODE: office,
      SD_SCHUL_CODE: school,
      MLSV_FROM_YMD: fromYmd,
      MLSV_TO_YMD: toYmd
    });
  }

  function rowKey(r) {
    if (!r) return "";
    return String(r.MLSV_YMD || "") + "|" + String(r.MMEAL_SC_CODE || "");
  }

  function isNoData(parsed) {
    return parsed.code === "INFO-200" || (!parsed.rows.length && !parsed.total);
  }

  async function fetchMealPage(fetchFn, office, school, from, to, page) {
    var url = mealQuery(office, school, from, to, page);
    var data = await fetchNeisJson(url, fetchFn);
    return readNeisPayload(data, "mealServiceDietInfo");
  }

  async function fetchMealsByDay(fetchFn, office, school, year, month) {
    var last = lastDayOfMonth(year, month);
    var days = [];
    var rows = [];
    var i = 0;
    var workers = [];
    var conc = 5;
    var w;
    function nextDay() {
      var day = days[i];
      i += 1;
      return day;
    }
    async function worker() {
      var day, ymd, page;
      while (i < days.length) {
        day = nextDay();
        ymd = year + pad2(month) + pad2(day);
        page = await fetchMealPage(fetchFn, office, school, ymd, ymd, 1);
        if (page.rows && page.rows.length) rows = rows.concat(page.rows);
      }
    }
    for (w = 1; w <= last; w++) days.push(w);
    for (w = 0; w < conc; w++) workers.push(worker());
    await Promise.all(workers);
    return rows;
  }

  async function searchNeisSchools(name, opts) {
    name = String(name || "").trim();
    if (!name) {
      throw new MealPlanError("학교 이름을 입력해 주세요.");
    }
    var fetchFn = getFetch(opts);
    var url = neisUrl("schoolInfo", {
      Type: "json",
      pIndex: 1,
      pSize: 100,
      SCHUL_NM: name
    });
    var data = await fetchNeisJson(url, fetchFn);
    var parsed = readNeisPayload(data, "schoolInfo");
    if (parsed.code === "INFO-200" || !parsed.rows.length) {
      throw new MealPlanError("검색한 이름의 학교를 찾지 못했습니다. 학교 이름을 다시 확인해 주세요.");
    }
    if (parsed.code && parsed.code !== "INFO-000") {
      throw new MealPlanError(parsed.message || "학교를 검색하지 못했습니다.");
    }
    var schools = parsed.rows.map(function (r) {
      return {
        SCHUL_NM: r.SCHUL_NM,
        ATPT_OFCDC_SC_CODE: r.ATPT_OFCDC_SC_CODE,
        SD_SCHUL_CODE: r.SD_SCHUL_CODE,
        ATPT_OFCDC_SC_NM: r.ATPT_OFCDC_SC_NM
      };
    });
    return {
      schools: schools,
      total: parsed.total || schools.length,
      truncated: (parsed.total || 0) > schools.length
    };
  }

  async function fetchNeisMeals(params) {
    params = params || {};
    var office = params.ATPT_OFCDC_SC_CODE || params.atptOfcdcScCode;
    var school = params.SD_SCHUL_CODE || params.sdSchulCode;
    var year = Number(params.year);
    var month = Number(params.month);
    var fetchFn = getFetch(params);
    var from;
    var to;
    var first;
    var rows;
    var total;
    var page;
    var next;
    var firstKey;
    if (!office || !school) {
      throw new MealPlanError("학교를 선택한 뒤 급식을 가져와 주세요.");
    }
    if (!year || month < 1 || month > 12) {
      throw new MealPlanError("연월을 선택해 주세요.");
    }
    from = year + pad2(month) + "01";
    to = year + pad2(month) + pad2(lastDayOfMonth(year, month));
    first = await fetchMealPage(fetchFn, office, school, from, to, 1);
    if (isNoData(first)) {
      throw new MealPlanError(
        "이 달의 급식 정보가 없습니다. 아직 나이스에 공개되지 않은 달일 수 있습니다."
      );
    }
    if (first.code && first.code !== "INFO-000") {
      throw new MealPlanError(first.message || "나이스 급식 정보를 가져오지 못했습니다.");
    }
    rows = first.rows.slice();
    total = first.total || rows.length;
    firstKey = rowKey(rows[0]);
    page = 2;
    while (rows.length < total) {
      next = await fetchMealPage(fetchFn, office, school, from, to, page);
      if (!next.rows.length) break;
      if (rowKey(next.rows[0]) === firstKey) break;
      rows = rows.concat(next.rows);
      if (next.rows.length < 100) break;
      page += 1;
      if (page > 50) break;
    }
    if (total && rows.length < total) {
      rows = await fetchMealsByDay(fetchFn, office, school, year, month);
    }
    return recordsFromNeisRows(rows);
  }

  function writeExcelJS(ExcelJS, records, options) {
    var wb = new ExcelJS.Workbook();
    var ws = wb.addWorksheet("Sheet1", {
      views: [{ state: "frozen", ySplit: 1 }]
    });
    ws.columns = [
      { width: 14 },
      { width: 15 },
      { width: 10 },
      { width: 62 },
      { width: 10 }
    ];
    var thin = { style: "thin", color: { argb: "FF999999" } };
    var border = { top: thin, left: thin, bottom: thin, right: thin };
    var font = { name: "Arial", size: 10 };
    var headerFont = { name: "Arial", size: 10, bold: true };
    var headerFill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFDDEBF7" }
    };

    var headerRow = ws.getRow(1);
    headerRow.height = 32;
    HEADERS.forEach(function (h, idx) {
      var cell = headerRow.getCell(idx + 1);
      cell.value = h;
      cell.font = headerFont;
      cell.alignment = { wrapText: true, horizontal: "center", vertical: "middle" };
      cell.fill = headerFill;
      cell.border = border;
    });

    records.forEach(function (rec, i) {
      var row = ws.getRow(i + 2);
      var a = row.getCell(1);
      a.value = toExcelSerial(rec.y, rec.m, rec.d);
      a.numFmt = "yyyymmdd";
      a.font = font;
      a.alignment = { horizontal: "center", vertical: "middle" };
      a.border = border;
      var b = row.getCell(2);
      b.value = rec.mealCode;
      b.font = font;
      b.alignment = { horizontal: "center", vertical: "middle" };
      b.border = border;
      var c = row.getCell(3);
      c.value = rec.mealName;
      c.font = font;
      c.alignment = { horizontal: "center", vertical: "middle" };
      c.border = border;
      var d = row.getCell(4);
      d.value = rec.menu;
      d.font = font;
      d.alignment = { wrapText: true, vertical: "middle" };
      d.border = border;
      var e = row.getCell(5);
      e.value = kcalForWrite(rec, options);
      e.font = font;
      e.alignment = { horizontal: "center", vertical: "middle" };
      e.border = border;
    });

    return wb.xlsx.writeBuffer();
  }

  return {
    MealPlanError: MealPlanError,
    parseMealPlan: parseMealPlan,
    parseWorkbook: parseWorkbook,
    pickMealSheet: pickMealSheet,
    writeSheetJS: writeSheetJS,
    writeExcelJS: writeExcelJS,
    kcalForWrite: kcalForWrite,
    toExcelSerial: toExcelSerial,
    parseKcal: parseKcal,
    parseNeisKcal: parseNeisKcal,
    dishesFromNeis: dishesFromNeis,
    recordsFromNeisRows: recordsFromNeisRows,
    recordFromNeisRow: recordFromNeisRow,
    readNeisPayload: readNeisPayload,
    searchNeisSchools: searchNeisSchools,
    fetchNeisMeals: fetchNeisMeals,
    neisUrl: neisUrl,
    DEFAULT_NEIS_SCHOOL: DEFAULT_NEIS_SCHOOL,
    HEADERS: HEADERS
  };
});
