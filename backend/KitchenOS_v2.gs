/**
 * =========================================================
 *  Sosisas · Kitchen Recipe OS — Apps Script extension (v2)
 * =========================================================
 *  Adds: team editing, inventory, stock moves, waste/yield log,
 *  counter tickets, daily reports, settings.
 *
 *  This file sits NEXT TO the existing Code.gs in the same
 *  Apps Script project. It does not touch the existing sheets
 *  (recipes / ingredients / feedbacks) or the existing actions
 *  (save_recipe / add_feedback). Two one-line hooks connect it:
 *
 *    // in doGet, just before returning the JSON object `result`:
 *    result = kos2_extendGet(result);
 *
 *    // at the top of doPost, right after parsing the payload:
 *    var kos2 = kos2_handlePost(payload);
 *    if (kos2) return kos2;
 *
 *  New sheets are created automatically on first use.
 *  Every write uses a script lock, so two phones writing at the
 *  same moment cannot corrupt a row.
 * =========================================================
 */

var KOS2 = {
  version: 2,
  // Optional: if this script is NOT bound to the spreadsheet, put its ID here.
  spreadsheetId: "",
  tables: {
    inventory_items: ["id", "name", "category", "unit", "min_qty", "aliases", "note", "created_at", "updated_at", "updated_by"],
    stock_moves:     ["id", "at", "day", "item_id", "item_name", "type", "qty", "unit", "delta", "counted", "section", "ref", "note", "chef", "created_at", "updated_at", "updated_by"],
    waste_log:       ["id", "at", "day", "item_name", "item_id", "process", "gross", "net", "unit", "yield_pct", "waste_qty", "reason", "note", "chef", "created_at", "updated_at", "updated_by"],
    tickets:         ["id", "no", "day", "created_at", "status", "done_at", "items", "note", "created_by", "done_by", "updated_at", "updated_by"],
    daily_reports:   ["id", "day", "total", "done", "void", "open_left", "avg_prep_min", "items_total", "breakdown", "closed_at", "closed_by", "created_at", "updated_at", "updated_by"],
    settings:        ["id", "value", "updated_at", "updated_by"]
  },
  jsonFields: { aliases: true, items: true, breakdown: true, value: true },
  textFields: { id: true, day: true, item_id: true, ref: true, no: false } // stored as plain text so Sheets never turns them into dates
};

/* ---------------- hooks ---------------- */

function kos2_extendGet(result) {
  result = result || {};
  try {
    result.kos = { version: KOS2.version };
    Object.keys(KOS2.tables).forEach(function (t) { result[t] = kos2_read(t); });
  } catch (e) {
    result.kos = { version: KOS2.version, error: String(e) };
  }
  return result;
}

function kos2_handlePost(payload) {
  if (!payload || payload.action !== "kos_batch") return null;
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var ops = Array.isArray(payload.ops) ? payload.ops : [];
    var applied = 0;
    ops.forEach(function (op) {
      if (op.op === "upsert" && KOS2.tables[op.table] && op.row && op.row.id) { kos2_upsert(op.table, [op.row]); applied++; }
      else if (op.op === "delete" && KOS2.tables[op.table] && op.id) { kos2_delete(op.table, [op.id]); applied++; }
      else if (op.op === "team" && Array.isArray(op.rows)) { kos2_saveTeam(op.rows); applied++; }
    });
    SpreadsheetApp.flush();
    return kos2_json({ status: "success", applied: applied });
  } catch (e) {
    return kos2_json({ status: "error", message: String(e && e.message || e) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- sheet helpers ---------------- */

function kos2_ss() {
  return KOS2.spreadsheetId ? SpreadsheetApp.openById(KOS2.spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
}

function kos2_json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Get (or create) a sheet and make sure every expected header exists. Returns {sheet, headers}. */
function kos2_sheet(name) {
  var headers = KOS2.tables[name];
  var ss = kos2_ss();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat("@"); // plain text: no auto-dates
    return { sheet: sh, headers: headers.slice() };
  }
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var cur = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  if (cur.length === 1 && cur[0] === "") {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    return { sheet: sh, headers: headers.slice() };
  }
  var missing = headers.filter(function (h) { return cur.indexOf(h) < 0; });
  if (missing.length) {
    sh.getRange(1, cur.length + 1, 1, missing.length).setValues([missing]).setFontWeight("bold");
    sh.getRange(1, cur.length + 1, sh.getMaxRows(), missing.length).setNumberFormat("@");
    cur = cur.concat(missing);
  }
  return { sheet: sh, headers: cur };
}

function kos2_cellOut(h, v) {
  if (v instanceof Date) {
    if (h === "day") return Utilities.formatDate(v, kos2_ss().getSpreadsheetTimeZone(), "yyyy-MM-dd");
    return v.toISOString();
  }
  if (KOS2.jsonFields[h] && typeof v === "string" && v !== "") {
    try { return JSON.parse(v); } catch (e) { return v; }
  }
  return v;
}

function kos2_cellIn(h, v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") return JSON.stringify(v);
  return v;
}

function kos2_read(name) {
  var s = kos2_sheet(name);
  var sh = s.sheet;
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, s.headers.length).getValues();
  var out = [];
  vals.forEach(function (row) {
    if (row.join("") === "") return;
    var o = {};
    s.headers.forEach(function (h, i) { if (h) o[h] = kos2_cellOut(h, row[i]); });
    o.id = String(o.id || "");
    if (o.id) out.push(o);
  });
  return out;
}

function kos2_upsert(name, rows) {
  var s = kos2_sheet(name);
  var sh = s.sheet, headers = s.headers;
  var idCol = headers.indexOf("id");
  var last = sh.getLastRow();
  var ids = last >= 2 ? sh.getRange(2, idCol + 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  rows.forEach(function (row) {
    var values = headers.map(function (h) { return kos2_cellIn(h, row[h]); });
    var i = ids.indexOf(String(row.id));
    if (i > -1) {
      sh.getRange(i + 2, 1, 1, headers.length).setValues([values]);
    } else {
      sh.getRange(sh.getLastRow() + 1, 1, 1, headers.length).setNumberFormat("@").setValues([values]);
      ids.push(String(row.id));
    }
  });
}

function kos2_delete(name, idList) {
  var s = kos2_sheet(name);
  var sh = s.sheet;
  var idCol = s.headers.indexOf("id");
  var last = sh.getLastRow();
  if (last < 2) return;
  var ids = sh.getRange(2, idCol + 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); });
  // delete bottom-up so row numbers stay valid
  for (var i = ids.length - 1; i >= 0; i--) {
    if (idList.indexOf(ids[i]) > -1) sh.deleteRow(i + 2);
  }
}

/* ---------------- team ---------------- */

/** Finds the existing team sheet: the one whose header row has chef_name AND station. */
function kos2_teamSheet() {
  var sheets = kos2_ss().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    if (sh.getLastColumn() < 1) continue;
    var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    if (h.indexOf("chef_name") > -1 && h.indexOf("station") > -1) return { sheet: sh, headers: h };
  }
  var sh2 = kos2_ss().insertSheet("team");
  sh2.getRange(1, 1, 1, 2).setValues([["chef_name", "station"]]).setFontWeight("bold");
  sh2.setFrozenRows(1);
  return { sheet: sh2, headers: ["chef_name", "station"] };
}

/** Replaces the team list, keeping any extra columns a chef already had (phone, role, …). */
function kos2_saveTeam(rows) {
  var t = kos2_teamSheet();
  var sh = t.sheet, headers = t.headers;
  var nameCol = headers.indexOf("chef_name");
  var last = sh.getLastRow();
  var old = {};
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, headers.length).getValues().forEach(function (r) { old[String(r[nameCol]).trim()] = r; });
  }
  var out = rows.map(function (r) {
    var base = old[String(r.chef_name).trim()] || headers.map(function () { return ""; });
    return headers.map(function (h, i) {
      if (h === "chef_name") return String(r.chef_name || "").trim();
      if (h === "station") return String(r.station || "").trim();
      return base[i];
    });
  });
  if (last >= 2) sh.getRange(2, 1, last - 1, headers.length).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, headers.length).setValues(out);
}

/* ---------------- manual test (run from the editor) ---------------- */
function kos2_selfTest() {
  var r = kos2_extendGet({});
  Logger.log(JSON.stringify({ version: r.kos, counts: Object.keys(KOS2.tables).map(function (t) { return t + ":" + r[t].length; }) }));
}
