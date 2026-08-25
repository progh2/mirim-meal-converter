"use strict";

var fs = require("fs");
var path = require("path");
var conv = require("./convert.js");

var SRC = process.env.MEAL_SRC ||
  "/workspace/uploads/6ffb32fb03e7e9f3dc51a3338ef54a0855634c22055899747e7756d392c7d321.xlsx";
var AUG = process.env.MEAL_AUG ||
  "/workspace/uploads/c76b58a9094680d32ca3dfb5e03c8c0f7fd85b8e682d3bef324b43d940748a0a.xlsx";
var XLSX_PATH = path.join(__dirname, "vendor", "xlsx.full.min.js");

var fails = [];
function assert(cond, msg) {
  if (!cond) fails.push(msg);
}

var NEIS_BREAKFAST = {
  ATPT_OFCDC_SC_CODE: "B10",
  SD_SCHUL_CODE: "7011569",
  SCHUL_NM: "미림마이스터고등학교",
  MMEAL_SC_CODE: "1",
  MMEAL_SC_NM: "조식",
  MLSV_YMD: "20250901",
  DDISH_NM: "추가밥 <br/>크로크무슈 (1.2.5.6.10)<br/>볶음김치 (9.13)<br/>소세지품은계란말이 (1.5.12)",
  CAL_INFO: "763.5 Kcal",
  NTR_INFO: "탄수화물(g) : 108.7<br/>단백질(g) : 27.9<br/>지방(g) : 21.1<br/>비타민A(R.E) : 143.4<br/>티아민(mg) : 0.3<br/>리보플라빈(mg) : 0.5<br/>비타민C(mg) : 3.8<br/>칼슘(mg) : 182.0<br/>철분(mg) : 3.3"
};

var NEIS_LUNCH_CALSU = {
  MMEAL_SC_CODE: "2",
  MMEAL_SC_NM: "중식",
  MLSV_YMD: "20250901",
  DDISH_NM: "쌀밥<br/>된장찌개<br/>",
  CAL_INFO: "1256.5 Kcal",
  NTR_INFO: "탄수화물(g) : 154.0<br/>단백질(g) : 51.5<br/>칼슔(mg) : 561.3<br/>철분(mg) : 5.1"
};

var NEIS_DINNER = {
  MMEAL_SC_CODE: "3",
  MMEAL_SC_NM: "석식",
  MLSV_YMD: "20250902",
  DDISH_NM: "추가밥 <br/>",
  CAL_INFO: "744.6 Kcal",
  NTR_INFO: "단백질(g) : 30.0<br/>칼슘(mg) : 100.0<br/>철분(mg) : 2.4"
};

var NEIS_AUG25_BREAKFAST = {
  ATPT_OFCDC_SC_CODE: "B10",
  SD_SCHUL_CODE: "7011569",
  SCHUL_NM: "미림마이스터고등학교",
  MMEAL_SC_CODE: "1",
  MMEAL_SC_NM: "조식",
  MLSV_YMD: "20260825",
  DDISH_NM: "보리쌀밥 <br/>단호박타락죽(j) (2.13)<br/>셀프토스트바(크로플) (1.2.5.6)<br/>청피망감자채볶음(j) (1.2.5.6.10.15.16)<br/>미니돈까스&케첩 (1.5.6.10.12)<br/>깍두기(조식) (9)",
  CAL_INFO: "948.4 Kcal",
  NTR_INFO: "탄수화물(g) : 100.0<br/>단백질(g) : 19.5<br/>지방(g) : 10.0<br/>칼슘(mg) : 119.3<br/>철분(mg) : 2.2"
};

console.log("NEIS mapping");

var kcalCalcium = conv.parseNeisKcal(NEIS_BREAKFAST.CAL_INFO, NEIS_BREAKFAST.NTR_INFO);
console.log("kcal 칼슘", kcalCalcium);
assert(kcalCalcium === 763.5, "NEIS kcal CAL_INFO only 763.5 got " + kcalCalcium);

var kcalCalsu = conv.parseNeisKcal(NEIS_LUNCH_CALSU.CAL_INFO, NEIS_LUNCH_CALSU.NTR_INFO);
console.log("kcal 칼슔", kcalCalsu);
assert(kcalCalsu === 1256.5, "NEIS kcal CAL_INFO only 1256.5 got " + kcalCalsu);

assert(conv.parseNeisKcal("1256.5 Kcal", "") === 1256.5, "CAL_INFO alone is Kcal");
assert(conv.parseNeisKcal("948.4 Kcal") === 948.4, "parseNeisKcal 948.4 Kcal");
var kcalSample = conv.parseKcal("708.6/19.1/245.8/4.5");
assert(kcalSample === 708.6, "parseKcal first number 708.6 got " + kcalSample);
assert(kcalSample !== 978, "must not sum excel nutrition parts");

var kcalAug = conv.parseNeisKcal(NEIS_AUG25_BREAKFAST.CAL_INFO, NEIS_AUG25_BREAKFAST.NTR_INFO);
console.log("kcal 2026-08-25 조식", kcalAug);
assert(kcalAug === 948.4, "parseNeisKcal 948.4 not 1089.4, got " + kcalAug);
assert(kcalAug !== 1089.4, "must not sum CAL_INFO + protein + calcium + iron");

var dishes = conv.dishesFromNeis(NEIS_BREAKFAST.DDISH_NM);
assert(dishes.length === 4, "split DDISH_NM on <br/>");
assert(dishes[0] === "추가밥", "trim first dish");
assert(dishes[1] === "크로크무슈1.2.5.6.10", "stick allergy digits to name");

var augDishes = conv.dishesFromNeis(NEIS_AUG25_BREAKFAST.DDISH_NM);
console.log("2026-08-25 dishes:", augDishes);
assert(augDishes.length === 6, "2026-08-25 breakfast has six lines, got " + augDishes.length);
assert(augDishes[0] === "보리쌀밥", "keep NEIS first dish 보리쌀밥");
assert(augDishes[1] === "단호박타락죽2.13", "drop (j), stick 2.13");
assert(augDishes[2] === "셀프토스트바(크로플)1.2.5.6", "keep name parens, stick allergy");
assert(augDishes[3] === "청피망감자채볶음1.2.5.6.10.15.16", "drop (j), keep NEIS order");
assert(augDishes[4] === "미니돈까스&케첩1.5.6.10.12", "stick allergy to 케첩");
assert(augDishes[5] === "깍두기9", "drop (조식), stick 9");
assert(augDishes.join("\n").indexOf("보리쌀밥\n단호박타락죽") === 0, "not only 보리쌀밥");
assert(conv.dishesFromNeis("깍두기(중식) (9)")[0] === "깍두기9", "drop (중식)");
assert(conv.dishesFromNeis("깍두기(석식) (9)")[0] === "깍두기9", "drop (석식)");

var augRec = conv.recordFromNeisRow(NEIS_AUG25_BREAKFAST);
assert(!!augRec, "recordFromNeisRow 2026-08-25 breakfast");
assert(augRec.menu.split("\n").length === 6, "record menu is full six lines");
assert(augRec.menu.indexOf("단호박타락죽2.13") !== -1, "record includes 단호박타락죽2.13");
assert(augRec.menu.indexOf("깍두기9") !== -1, "record includes 깍두기9");
assert(augRec.menu !== augRec.firstDish, "menu is not only firstDish");
assert(augRec.kcal === 948.4, "record kcal CAL_INFO 948.4 got " + augRec.kcal);

var mapped = conv.recordsFromNeisRows([NEIS_DINNER, NEIS_LUNCH_CALSU, NEIS_BREAKFAST, {
  MMEAL_SC_CODE: "2",
  MLSV_YMD: "20250903",
  DDISH_NM: "   <br/>  ",
  CAL_INFO: "10 Kcal",
  NTR_INFO: "단백질(g) : 1"
}]);
assert(mapped.filename === "mlsvTmplat_202509.xlsx", "filename " + mapped.filename);
assert(mapped.records.length === 3, "skip empty dishes, got " + mapped.records.length);
assert(mapped.records[0].mealCode === 0 && mapped.records[0].mealName === "조식", "MMEAL 1 → 조식 0");
assert(mapped.records[1].mealCode === 1 && mapped.records[1].mealName === "중식", "MMEAL 2 → 중식 1");
assert(mapped.records[2].mealCode === 2 && mapped.records[2].mealName === "석식", "MMEAL 3 → 석식 2");
assert(mapped.records[0].dateStr === "2025-09-01", "dateStr");
assert(mapped.records[0].menu.indexOf("\n") !== -1, "menu joined by newline");
assert(mapped.records[2].dateStr === "2025-09-02", "sort by date then meal");
assert(mapped.records[2].kcal === 744.6, "dinner CAL_INFO 744.6 got " + mapped.records[2].kcal);

var JOKE = conv.ZERO_KCAL_TEXT;
assert(JOKE === "맛있게 먹으면 0Kcal", "ZERO_KCAL_TEXT exact string with capital K");
assert(conv.kcalForWrite(mapped.records[0]) === 763.5, "kcalForWrite off keeps CAL_INFO");
assert(conv.kcalForWrite(mapped.records[0], {}) === 763.5, "kcalForWrite empty options keeps kcal");
assert(conv.kcalForWrite(mapped.records[0], { zeroKcal: false }) === 763.5, "kcalForWrite false keeps kcal");
assert(conv.kcalForWrite(mapped.records[0], { zeroKcal: true }) === JOKE, "kcalForWrite true writes joke string");
assert(conv.kcalForWrite(mapped.records[0], { zeroKcal: true }) !== 0, "kcalForWrite true is not numeric 0");
assert(typeof conv.kcalForWrite(mapped.records[0], { zeroKcal: true }) === "string", "kcalForWrite true is a string");
var menuBeforeZero = mapped.records[0].menu;
conv.kcalForWrite(mapped.records[0], { zeroKcal: true });
assert(mapped.records[0].menu === menuBeforeZero, "kcalForWrite does not rewrite menu");
mapped.records.forEach(function (rec) {
  var written = conv.kcalForWrite(rec, { zeroKcal: true });
  assert(written === JOKE, "every row writes joke string when flag on");
  assert(typeof written === "string", "every checked row Kcal is a string");
  assert(written !== 0, "checked Kcal is not numeric 0");
  assert(conv.kcalForWrite(rec, { zeroKcal: false }) === rec.kcal, "every row keeps kcal when flag off");
});

function writeAoa(records, options) {
  var aoa = null;
  var stub = {
    utils: {
      aoa_to_sheet: function (rows) {
        aoa = rows;
        return { "!ref": "A1:E1" };
      },
      decode_range: function () {
        return { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
      },
      encode_cell: function () { return "A1"; },
      book_new: function () { return {}; },
      book_append_sheet: function () {}
    },
    write: function () { return []; }
  };
  conv.writeSheetJS(stub, records, options);
  return aoa;
}

var aoaOff = writeAoa(mapped.records);
assert(aoaOff.length === mapped.records.length + 1, "writeSheetJS row count");
assert(aoaOff[1][4] === mapped.records[0].kcal, "writeSheetJS keeps kcal when flag off");
assert(aoaOff[1][3] === mapped.records[0].menu, "writeSheetJS keeps menu when flag off");

var aoaOn = writeAoa(mapped.records, { zeroKcal: true });
assert(aoaOn[1][3] === mapped.records[0].menu, "writeSheetJS zeroKcal does not write joke into menu");
aoaOn.slice(1).forEach(function (row, i) {
  assert(row[4] === JOKE, "writeSheetJS zeroKcal row " + (i + 1) + " Kcal joke string");
  assert(typeof row[4] === "string", "writeSheetJS zeroKcal row " + (i + 1) + " Kcal is a string");
  assert(row[4] !== 0, "writeSheetJS zeroKcal row " + (i + 1) + " Kcal is not numeric 0");
  assert(row[3] === mapped.records[i].menu, "writeSheetJS zeroKcal row " + (i + 1) + " menu unchanged");
});

var rec1393 = {
  y: 2026,
  m: 8,
  d: 25,
  dateStr: "2026-08-25",
  mealCode: 1,
  mealName: "중식",
  menu: "쌀밥\n된장찌개",
  kcal: 1393.1,
  firstDish: "쌀밥"
};
assert(conv.kcalForWrite(rec1393, { zeroKcal: true }) === JOKE, "kcalForWrite 1393.1 → joke string");
assert(conv.kcalForWrite(rec1393, { zeroKcal: false }) === 1393.1, "kcalForWrite false keeps 1393.1");
var copiedZero = conv.recordsForWrite([rec1393], { zeroKcal: true });
assert(copiedZero[0].kcal === JOKE, "recordsForWrite sets joke string");
assert(typeof copiedZero[0].kcal === "string", "recordsForWrite kcal is a string");
assert(copiedZero[0] !== rec1393, "recordsForWrite returns a shallow copy");
assert(rec1393.kcal === 1393.1, "recordsForWrite does not mutate source kcal");
assert(copiedZero[0].menu === rec1393.menu, "recordsForWrite keeps menu");
assert(conv.recordsForWrite([rec1393], { zeroKcal: false })[0].kcal === 1393.1, "recordsForWrite off keeps 1393.1");

var aoa1393Off = writeAoa([rec1393], { zeroKcal: false });
assert(aoa1393Off[1][4] === 1393.1, "writeSheetJS false keeps 1393.1");
var aoa1393On = writeAoa([rec1393], { zeroKcal: true });
assert(aoa1393On[1][4] === JOKE, "writeSheetJS true writes joke string not 1393.1");
assert(typeof aoa1393On[1][4] === "string", "writeSheetJS true Kcal is a string");
assert(aoa1393On[1][4] !== 0, "writeSheetJS true Kcal is not numeric 0");
assert(aoa1393On[1][3] === rec1393.menu, "writeSheetJS true does not write joke into menu");
assert(JSON.stringify(aoa1393On).indexOf("1393.1") === -1, "zeroKcal aoa cannot emit 1393.1");
assert(JSON.stringify(aoa1393On).indexOf(JOKE) !== -1, "zeroKcal aoa emits joke string");
assert(rec1393.kcal === 1393.1, "writeSheetJS does not mutate source 1393.1");

var alreadyJoke = conv.recordsForWrite([rec1393], { zeroKcal: true });
var aoaPreJoke = writeAoa(alreadyJoke);
assert(aoaPreJoke[1][4] === JOKE, "pre-mapped records write joke string even without options");
assert(typeof aoaPreJoke[1][4] === "string", "pre-mapped records write a string");

function writeExcelValues(records, options) {
  var values = [];
  var ExcelJS = {
    Workbook: function () {
      this.addWorksheet = function () {
        return {
          columns: null,
          getRow: function (n) {
            return {
              height: 0,
              getCell: function (col) {
                var cell = {
                  font: null,
                  alignment: null,
                  border: null,
                  fill: null,
                  numFmt: null
                };
                Object.defineProperty(cell, "value", {
                  set: function (v) {
                    if (n > 1) {
                      values[n - 2] = values[n - 2] || [];
                      values[n - 2][col - 1] = v;
                    }
                  },
                  get: function () {
                    return n > 1 && values[n - 2] ? values[n - 2][col - 1] : null;
                  }
                });
                return cell;
              }
            };
          }
        };
      };
      this.xlsx = {
        writeBuffer: function () {
          return Promise.resolve([]);
        }
      };
    }
  };
  return conv.writeExcelJS(ExcelJS, records, options).then(function () {
    return values;
  });
}

var html = fs.readFileSync(path.join(__dirname, "index.html"), "utf8");
assert(/styles\.css\?v=/.test(html), "styles.css must have cache-bust query");
assert(/convert\.js\?v=/.test(html), "convert.js must have cache-bust query");
assert(/app\.js\?v=/.test(html), "app.js must have cache-bust query");
var appSrc = fs.readFileSync(path.join(__dirname, "app.js"), "utf8");
assert(/recordsForDownload/.test(appSrc), "download maps a shallow copy before write");
assert(/copy\.kcal\s*=\s*text/.test(appSrc) || /copy\.kcal\s*=\s*zeroKcalText/.test(appSrc) || /copy\.kcal\s*=/.test(appSrc), "download copy sets joke kcal when checked");
assert(/writeSheetJS\(XLSX,\s*records,\s*options\)/.test(appSrc), "download still passes options into writeSheetJS");
assert(appSrc.indexOf("맛있게 먹으면 0Kcal") !== -1, "preview/download uses exact joke string");
assert(appSrc.indexOf("escapeHtml(zero ? 0") === -1, "preview must not show numeric 0 when checked");
assert(/zero \? kcalText/.test(appSrc) || /zero \? zeroKcalText/.test(appSrc), "preview Kcal uses joke string when checked");

try {
  conv.recordsFromNeisRows([]);
  fails.push("empty NEIS rows should throw");
} catch (e) {
  assert(e.name === "MealPlanError", "empty rows MealPlanError");
  assert(e.message.indexOf("급식 정보") !== -1, "empty month Korean error");
}

function jsonRes(obj) {
  return {
    ok: true,
    json: function () { return Promise.resolve(obj); }
  };
}

function schoolPayload(rows, total) {
  return {
    schoolInfo: [
      { head: [{ list_total_count: total == null ? rows.length : total }, { RESULT: { CODE: "INFO-000", MESSAGE: "ok" } }] },
      { row: rows }
    ]
  };
}

function mealPayload(rows, total) {
  return {
    mealServiceDietInfo: [
      { head: [{ list_total_count: total == null ? rows.length : total }, { RESULT: { CODE: "INFO-000", MESSAGE: "ok" } }] },
      { row: rows }
    ]
  };
}

(async function () {
  var search = await conv.searchNeisSchools("미림", {
    fetch: function (url) {
      assert(url.indexOf("schoolInfo") !== -1, "schoolInfo endpoint");
      assert(url.indexOf("Type=json") !== -1, "school Type=json");
      assert(url.indexOf("SCHUL_NM=") !== -1, "SCHUL_NM");
      return Promise.resolve(jsonRes(schoolPayload([
        { SCHUL_NM: "미림마이스터고등학교", ATPT_OFCDC_SC_CODE: "B10", SD_SCHUL_CODE: "7011569", ATPT_OFCDC_SC_NM: "서울특별시교육청" },
        { SCHUL_NM: "미림여자고등학교", ATPT_OFCDC_SC_CODE: "B10", SD_SCHUL_CODE: "7010167", ATPT_OFCDC_SC_NM: "서울특별시교육청" }
      ])));
    }
  });
  assert(search.schools.length === 2, "school pick list");
  assert(search.schools[0].SD_SCHUL_CODE === "7011569", "school code");

  var calls = [];
  var monthRows = [
    NEIS_BREAKFAST,
    NEIS_LUNCH_CALSU,
    NEIS_DINNER,
    {
      MMEAL_SC_CODE: "1",
      MLSV_YMD: "20250903",
      DDISH_NM: "죽",
      CAL_INFO: "100 Kcal",
      NTR_INFO: "단백질(g) : 10<br/>칼슘(mg) : 20<br/>철분(mg) : 1"
    }
  ];
  var fetched = await conv.fetchNeisMeals({
    ATPT_OFCDC_SC_CODE: "B10",
    SD_SCHUL_CODE: "7011569",
    year: 2025,
    month: 9,
    fetch: function (url) {
      calls.push(url);
      assert(url.indexOf("mealServiceDietInfo") !== -1, "meal endpoint");
      assert(url.indexOf("pSize=100") !== -1, "pSize=100");
      assert(url.indexOf("MLSV_FROM_YMD=") !== -1 && url.indexOf("MLSV_TO_YMD=") !== -1, "month range");
      if (url.indexOf("pIndex=1") !== -1 && url.indexOf("MLSV_FROM_YMD=20250901") !== -1 && url.indexOf("MLSV_TO_YMD=20250930") !== -1) {
        return Promise.resolve(jsonRes(mealPayload(monthRows.slice(0, 1), 4)));
      }
      if (url.indexOf("pIndex=2") !== -1) {
        return Promise.resolve(jsonRes(mealPayload(monthRows.slice(0, 1), 4)));
      }
      var m = url.match(/MLSV_FROM_YMD=(\d{8})/);
      var day = m && m[1];
      var dayRows = monthRows.filter(function (r) { return r.MLSV_YMD === day; });
      if (!dayRows.length) {
        return Promise.resolve(jsonRes({ RESULT: { CODE: "INFO-200", MESSAGE: "해당하는 데이터가 없습니다." } }));
      }
      return Promise.resolve(jsonRes(mealPayload(dayRows, dayRows.length)));
    }
  });
  assert(fetched.records.length === 4, "daily fallback assembled month, got " + fetched.records.length);
  assert(calls.some(function (u) { return /MLSV_FROM_YMD=20250901/.test(u) && /MLSV_TO_YMD=20250901/.test(u); }), "fallback uses same-day range");
  assert(fetched.records[0].kcal === 763.5, "fetched breakfast kcal CAL_INFO only");

  try {
    await conv.fetchNeisMeals({
      ATPT_OFCDC_SC_CODE: "B10",
      SD_SCHUL_CODE: "7011569",
      year: 2026,
      month: 12,
      fetch: function () {
        return Promise.resolve(jsonRes({ RESULT: { CODE: "INFO-200", MESSAGE: "해당하는 데이터가 없습니다." } }));
      }
    });
    fails.push("empty future month should throw");
  } catch (e) {
    assert(e.name === "MealPlanError", "empty month MealPlanError");
    assert(e.message.indexOf("공개되지") !== -1, "unpublished month message");
  }

  if (fs.existsSync(SRC) && fs.existsSync(XLSX_PATH)) {
    runExcelTests(require(XLSX_PATH));
  } else {
    console.log("skip excel fixtures (set MEAL_SRC and vendor/xlsx.full.min.js to run)");
  }

  var excelOn = await writeExcelValues([rec1393], { zeroKcal: true });
  assert(excelOn[0][4] === JOKE, "writeExcelJS true writes joke string not 1393.1");
  assert(typeof excelOn[0][4] === "string", "writeExcelJS true Kcal is a string");
  assert(excelOn[0][4] !== 0, "writeExcelJS true Kcal is not numeric 0");
  assert(excelOn[0][3] === rec1393.menu, "writeExcelJS true keeps menu");
  var excelOff = await writeExcelValues([rec1393], { zeroKcal: false });
  assert(excelOff[0][4] === 1393.1, "writeExcelJS false keeps 1393.1");

  if (fails.length) {
    console.error("\nFAIL (" + fails.length + ")");
    fails.forEach(function (f) { console.error(" -", f); });
    process.exit(1);
  }
  console.log("\nPASS neis mapping + kcal");
})().catch(function (err) {
  console.error(err);
  process.exit(1);
});

function runExcelTests(XLSX) {
  var buf = fs.readFileSync(SRC);
  var wb = XLSX.read(buf, { type: "buffer", cellDates: true });
  console.log("sheets:", wb.SheetNames);

  var result = conv.parseWorkbook(XLSX, wb);
  var records = result.records;

  console.log("sheet:", result.sheetName);
  console.log("filename:", result.filename);
  console.log("row count:", records.length);
  console.log("month:", result.year, result.month);

  function find(dateStr, meal) {
    return records.find(function (r) {
      return r.dateStr === dateStr && r.mealName === meal;
    });
  }

  var b0901 = find("2026-09-01", "조식");
  console.log("9/1 조식:", b0901 ? ("kcal=" + b0901.kcal + " first=" + b0901.firstDish) : "MISSING");

  assert(records.length >= 54 && records.length <= 58, "expected ~56 rows, got " + records.length);
  assert(records.length === 56, "exact 56 rows preferred, got " + records.length);
  assert(!!b0901, "missing 9/1 조식");
  assert(b0901 && b0901.kcal === 708.6, "9/1 조식 kcal expected 708.6 got " + (b0901 && b0901.kcal));
  assert(b0901 && b0901.firstDish.indexOf("루테인쌀밥") !== -1, "9/1 조식 first dish");

  assert(!find("2026-09-04", "석식"), "should skip 9/4 석식");
  assert(!!find("2026-09-04", "조식"), "should keep 9/4 조식");
  assert(!!find("2026-09-04", "중식"), "should keep 9/4 중식");

  assert(!find("2026-09-24", "조식") && !find("2026-09-24", "중식") && !find("2026-09-24", "석식"), "no 9/24");
  assert(!find("2026-09-25", "조식") && !find("2026-09-25", "중식") && !find("2026-09-25", "석식"), "no 9/25");

  var weekend = records.filter(function (r) {
    var day = new Date(r.y, r.m - 1, r.d).getDay();
    return day === 0 || day === 6;
  });
  assert(weekend.length === 0, "weekend rows: " + weekend.map(function (r) { return r.dateStr; }).join(","));

  var d0914 = find("2026-09-14", "석식");
  console.log("9/14 석식 first:", d0914 && d0914.firstDish);
  assert(d0914 && d0914.firstDish.indexOf("등심김치카츠나베") !== -1, "week3 석식 overflow first dish");

  assert(!!find("2026-09-30", "조식"), "9/30 text date 조식");
  assert(!!find("2026-09-30", "중식"), "9/30 중식");
  assert(!!find("2026-09-30", "석식"), "9/30 석식");

  assert(!find("2026-09-23", "석식"), "no 9/23 석식");
  assert(!find("2026-09-18", "석식"), "no 9/18 석식");
  assert(!find("2026-09-11", "석식"), "no 9/11 석식");

  var kcalExcel = conv.parseKcal("708.6/19.1/245.8/4.5");
  assert(kcalExcel === 708.6, "parseKcal first number 708.6 got " + kcalExcel);
  assert(kcalExcel !== 978, "must not sum excel nutrition parts");

  records.forEach(function (r) {
    console.log(r.dateStr, r.mealName, "kcal=" + r.kcal, "|", r.firstDish);
  });

  var outDir = path.join(__dirname, "test-output");
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);
  var outBuf = Buffer.from(conv.writeSheetJS(XLSX, records));
  var outPath = path.join(outDir, result.filename);
  fs.writeFileSync(outPath, outBuf);
  console.log("wrote", outPath, outBuf.length, "bytes");

  var wb2 = XLSX.read(outBuf, { type: "buffer", cellDates: true });
  var grid2 = XLSX.utils.sheet_to_json(wb2.Sheets.Sheet1, { header: 1, raw: true, defval: null });
  assert(grid2.length === records.length + 1, "output rows");
  var a2 = grid2[1][0];
  console.log("output A2", a2, typeof a2);
  if (a2 instanceof Date) {
    assert(a2.getUTCFullYear() === 2026 && a2.getUTCMonth() === 8 && a2.getUTCDate() === 1, "A2 date 2026-09-01");
  } else if (typeof a2 === "number") {
    var serial = conv.toExcelSerial(2026, 9, 1);
    assert(a2 === serial, "A2 serial " + a2 + " expected " + serial);
  }
  assert(grid2[1][4] === records[0].kcal, "writeSheetJS keeps kcal when zeroKcal off");
  assert(grid2[1][3] === records[0].menu, "writeSheetJS keeps menu when zeroKcal off");

  var zeroBuf = Buffer.from(conv.writeSheetJS(XLSX, records, { zeroKcal: true }));
  var gridZero = XLSX.utils.sheet_to_json(
    XLSX.read(zeroBuf, { type: "buffer", cellDates: true }).Sheets.Sheet1,
    { header: 1, raw: true, defval: null }
  );
  assert(gridZero.length === records.length + 1, "zeroKcal output rows");
  assert(gridZero[1][3] === records[0].menu, "zeroKcal must not change menu cell");
  var zi;
  for (zi = 1; zi < gridZero.length; zi++) {
    assert(gridZero[zi][4] === JOKE, "zeroKcal row " + zi + " Kcal expected joke string got " + gridZero[zi][4]);
    assert(typeof gridZero[zi][4] === "string", "zeroKcal row " + zi + " Kcal is a string");
    assert(gridZero[zi][4] !== 0, "zeroKcal row " + zi + " Kcal is not numeric 0");
    assert(gridZero[zi][3] === records[zi - 1].menu, "zeroKcal row " + zi + " menu unchanged");
  }
  var rec1393Excel = {
    y: 2026, m: 8, d: 25, dateStr: "2026-08-25",
    mealCode: 1, mealName: "중식", menu: "쌀밥", kcal: 1393.1, firstDish: "쌀밥"
  };
  var buf1393 = Buffer.from(conv.writeSheetJS(XLSX, [rec1393Excel], { zeroKcal: true }));
  var grid1393 = XLSX.utils.sheet_to_json(
    XLSX.read(buf1393, { type: "buffer", cellDates: true }).Sheets.Sheet1,
    { header: 1, raw: true, defval: null }
  );
  assert(grid1393[1][4] === JOKE, "xlsx zeroKcal writes joke string, got " + grid1393[1][4]);
  assert(typeof grid1393[1][4] === "string", "xlsx zeroKcal writes a string");
  assert(grid1393[1][4] !== 0, "xlsx zeroKcal is not numeric 0");
  assert(buf1393.indexOf(Buffer.from("1393.1")) === -1, "xlsx bytes must not contain 1393.1");
  assert(buf1393.indexOf(Buffer.from(JOKE, "utf8")) !== -1, "xlsx bytes contain joke string");
  var buf1393Off = Buffer.from(conv.writeSheetJS(XLSX, [rec1393Excel], { zeroKcal: false }));
  var grid1393Off = XLSX.utils.sheet_to_json(
    XLSX.read(buf1393Off, { type: "buffer", cellDates: true }).Sheets.Sheet1,
    { header: 1, raw: true, defval: null }
  );
  assert(grid1393Off[1][4] === 1393.1, "xlsx unchecked keeps 1393.1");

  if (fs.existsSync(AUG)) {
    try {
      var wbAug = XLSX.read(fs.readFileSync(AUG), { type: "buffer", cellDates: true });
      conv.parseWorkbook(XLSX, wbAug);
      fails.push("August template should be rejected");
    } catch (e) {
      console.log("non-meal rejection:", e.message);
      assert(!!e.message && e.message.indexOf("식단") !== -1, "Korean error for non-meal file");
    }
  }
}
