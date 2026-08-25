"use strict";

var fs = require("fs");
var path = require("path");
var XLSX = require("./vendor/xlsx.full.min.js");
var conv = require("./convert.js");

var SRC = process.env.MEAL_SRC ||
  "/workspace/uploads/6ffb32fb03e7e9f3dc51a3338ef54a0855634c22055899747e7756d392c7d321.xlsx";

var fails = [];
function assert(cond, msg) {
  if (!cond) fails.push(msg);
}

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
assert(b0901 && b0901.kcal === 978, "9/1 조식 kcal expected 978 got " + (b0901 && b0901.kcal));
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

var kcalSample = conv.parseKcal("708.6/19.1/245.8/4.5");
assert(kcalSample === 978, "parseKcal sum 978 got " + kcalSample);

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


var AUG = "/workspace/uploads/c76b58a9094680d32ca3dfb5e03c8c0f7fd85b8e682d3bef324b43d940748a0a.xlsx";
try {
  var wbAug = XLSX.read(fs.readFileSync(AUG), { type: "buffer", cellDates: true });
  conv.parseWorkbook(XLSX, wbAug);
  fails.push("August template should be rejected");
} catch (e) {
  console.log("non-meal rejection:", e.message);
  assert(!!e.message && e.message.indexOf("식단") !== -1, "Korean error for non-meal file");
}

if (fails.length) {
  console.error("\nFAIL (" + fails.length + ")");
  fails.forEach(function (f) { console.error(" -", f); });
  process.exit(1);
}
console.log("\nPASS rows=" + records.length + " kcal_9/1_조식=" + b0901.kcal);
