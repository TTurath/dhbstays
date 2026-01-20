/**
 * File: 00_Core.gs
 * DHB STAYS PMS v2 — Core Foundation (Clean Build)
 *
 * ✅ This is the base layer that ALL other files depend on:
 *   - constants (sheet names, roles)
 *   - safe helpers (dates, parsing, json)
 *   - sheet helpers (get, ensure, indexMap)
 *   - settings reader
 *   - country codes loader + phone normalize + country lookup
 *   - audit logger (safe)
 *
 * IMPORTANT:
 * - Avoid modern syntax that may break in some Apps Script runtimes.
 * - No ?? , no optional chaining, no fancy assignments.
 */

/* =========================
 * GLOBAL CONFIG
 * ========================= */

var DHB = {
  TZ: "Africa/Cairo",
  LOCALE: "ar_EG",

  ROLES: ["Admin", "FrontOffice", "Housekeeping", "Security", "Viewer"],

  SHEETS: {
    README: "README",
    SETTINGS: "Settings",
    USERS: "Users",
    UNITS: "Units",
    RES: "Reservations",
    GUESTS: "Guests",
    PRICES: "Prices",
    MAINT: "MaintenanceLog",
    LISTINGS: "Listings",
    COUNTRY: "CountryCodes",
    CLEAN_REQ: "CleaningRequests",
    CHECKLIST: "UnitDailyChecklist",
    STATUS_HIST: "UnitStatusHistory",
    RAW: "Imports_Raw",
    ERRORS: "Import_Errors",
    AUDIT: "AuditLog"
  },

  // Import alias keys (for CSV headers)
  IMPORT_ALIASES: {
    RES_ID: ["confirmation code", "res_id", "reservation id", "confirmation", "code"],
    STATUS: ["status", "reservation status"],
    GUEST_NAME: ["guest name", "guest", "name"],
    PHONE: ["contact", "phone", "phone_raw", "guest phone", "telephone"],
    ADULTS: ["# of adults", "adults", "adult", "num_adults"],
    CHILDREN: ["# of children", "children", "child", "num_children"],
    INFANTS: ["# of infants", "infants", "infant", "num_infants"],
    GUESTS_TOTAL: ["guests_total", "total guests", "guests", "guest_count", "total_guests"],
    START: ["start date", "check_in", "check in", "arrival"],
    END: ["end date", "check_out", "check out", "departure"],
    NIGHTS: ["# of nights", "nights", "night"],
    BOOKED_AT: ["booked", "booked_at", "booking date", "created"],
    LISTING: ["listing", "listing_name", "property", "unit"],
    EARNINGS: ["earnings", "earnings_usd", "payout", "amount", "total payout"],
    CHANNEL: ["channel", "source", "platform"]
  }
};

/* =========================
 * CORE: Date Helpers
 * ========================= */

function dhbNormalizeToMidnight_(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return d;
  d.setHours(0, 0, 0, 0);
  return d;
}

function dhbToISODate_(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return "";
  var y = d.getFullYear();
  var m = String(d.getMonth() + 1);
  if (m.length < 2) m = "0" + m;
  var day = String(d.getDate());
  if (day.length < 2) day = "0" + day;
  return y + "-" + m + "-" + day;
}

function dhbNow_() {
  return new Date();
}

/* =========================
 * CORE: Safe parsing
 * ========================= */

function dhbStr_(v) {
  return String(v === null || v === undefined ? "" : v).trim();
}

function dhbToInt_(v, fallback) {
  if (v === null || v === undefined || v === "") return Number(fallback || 0);
  if (typeof v === "number" && isFinite(v)) return Math.floor(v);
  var s = String(v).replace(/[^\d\-]+/g, "");
  var n = parseInt(s, 10);
  return isNaN(n) ? Number(fallback || 0) : n;
}

function dhbParseMoney_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number" && isFinite(v)) return v;
  var s = String(v).trim();
  if (!s) return 0;
  var cleaned = s.replace(/[^0-9.\-]/g, "");
  var num = Number(cleaned);
  return isNaN(num) ? 0 : num;
}

/**
 * Loose date parse:
 * - Date object => keep
 * - YYYY-MM-DD
 * - DD/MM/YYYY or MM/DD/YYYY (best guess)
 */
function dhbParseDateLoose_(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return dhbNormalizeToMidnight_(v);

  var s = String(v).trim();
  if (!s) return null;

  var iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) return dhbNormalizeToMidnight_(new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])));

  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    var a = Number(m[1]), b = Number(m[2]), y = Number(m[3]);
    // If first part > 12 => dd/mm/yyyy
    if (a > 12) return dhbNormalizeToMidnight_(new Date(y, b - 1, a));
    // Else treat as mm/dd/yyyy
    return dhbNormalizeToMidnight_(new Date(y, a - 1, b));
  }

  var d = new Date(s);
  if (isNaN(d.getTime())) return null;
  return dhbNormalizeToMidnight_(d);
}

function dhbSafeJson_(obj) {
  try {
    if (obj === undefined) return "";
    return JSON.stringify(obj);
  } catch (e) {
    return String(obj);
  }
}

/* =========================
 * CORE: Sheet helpers
 * ========================= */

function dhbSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function dhbGetSheetOrThrow_(name) {
  var ss = dhbSS_();
  var sh = ss.getSheetByName(String(name || "").trim());
  if (!sh) throw new Error("Missing sheet: " + name);
  return sh;
}

function dhbGetSheetOrNull_(name) {
  var ss = dhbSS_();
  return ss.getSheetByName(String(name || "").trim());
}

function dhbEnsureSheet_(name) {
  var ss = dhbSS_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/** strict index map (exact header text) */
function dhbIndexMapStrict_(headerRow) {
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var key = String(headerRow[i] || "").trim();
    if (key) map[key] = i;
  }
  return map;
}

/** normalize header keys to "human" comparable form */
function dhbNormHeaderHuman_(h) {
  return String(h || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[^\w #]+/g, "")
    .trim();
}

/** build alias map: normalizedKey -> originalKey */
function dhbBuildAliasMapFromRowObject_(obj) {
  var out = {};
  var keys = Object.keys(obj || {});
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (k === "__rowIndex") continue;
    var nk = dhbNormHeaderHuman_(k);
    if (nk && !out[nk]) out[nk] = k;
  }
  return out;
}

function dhbPickByAliases_(rowObj, aliasMap, wantedList) {
  for (var i = 0; i < wantedList.length; i++) {
    var nk = String(wantedList[i] || "").trim().toLowerCase();
    var origKey = aliasMap[nk];
    if (origKey && Object.prototype.hasOwnProperty.call(rowObj, origKey)) {
      return rowObj[origKey];
    }
  }
  return "";
}

/* =========================
 * CORE: Settings
 * ========================= */

function dhbGetSetting_(key, fallback) {
  try {
    var sh = dhbGetSheetOrNull_(DHB.SHEETS.SETTINGS);
    if (!sh) return fallback || "";
    var v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][0] || "").trim() === key) {
        var val = String(v[i][1] || "").trim();
        return val || (fallback || "");
      }
    }
  } catch (e) {}
  return fallback || "";
}

/* =========================
 * CORE: Audit log (safe)
 * ========================= */

function dhbAuditSafe_(username, action, detailsObj) {
  try {
    var sh = dhbGetSheetOrNull_(DHB.SHEETS.AUDIT);
    if (!sh) return;
    sh.appendRow([new Date(), String(username || "SYSTEM"), String(action || ""), dhbSafeJson_(detailsObj)]);
  } catch (e) {
    // swallow
  }
}

/* =========================
 * CORE: Phone + Country Codes (v2)
 * ========================= */

function dhbToLatinDigits_(s) {
  var map = {
    "٠": "0", "١": "1", "٢": "2", "٣": "3", "٤": "4", "٥": "5", "٦": "6", "٧": "7", "٨": "8", "٩": "9",
    "۰": "0", "۱": "1", "۲": "2", "۳": "3", "۴": "4", "۵": "5", "۶": "6", "۷": "7", "۸": "8", "۹": "9"
  };
  return String(s || "").replace(/[٠-٩۰-۹]/g, function (ch) { return map[ch] || ch; });
}

/**
 * Loads CountryCodes sheet list:
 * headers: calling_code | iso2 | name_ar | name_en | flag
 * returns: [{calling, iso2, name_ar, name_en, flag}]
 */
function dhbLoadCountryCodes_() {
  var sh = dhbGetSheetOrNull_(DHB.SHEETS.COUNTRY);
  if (!sh) return [];
  var v = sh.getDataRange().getValues();
  if (!v || v.length < 2) return [];

  var idx = dhbIndexMapStrict_(v[0]);
  var callingIx = idx.calling_code;
  var isoIx = idx.iso2;
  var arIx = idx.name_ar;
  var enIx = idx.name_en;
  var flagIx = idx.flag;

  // fallback positions if headers changed
  if (callingIx === undefined) callingIx = 0;
  if (isoIx === undefined) isoIx = 1;
  if (arIx === undefined) arIx = 2;
  if (enIx === undefined) enIx = 3;
  if (flagIx === undefined) flagIx = 4;

  var out = [];
  for (var i = 1; i < v.length; i++) {
    var calling = String(v[i][callingIx] || "").replace(/\D/g, "").trim();
    if (!calling) continue;
    out.push({
      calling: calling,
      iso2: String(v[i][isoIx] || "").trim(),
      name_ar: String(v[i][arIx] || "").trim(),
      name_en: String(v[i][enIx] || "").trim(),
      flag: String(v[i][flagIx] || "🌍").trim() || "🌍"
    });
  }

  // longest calling code first (important for matching)
  out.sort(function (a, b) { return String(b.calling).length - String(a.calling).length; });
  return out;
}

function dhbNormalizePhone_(phoneRaw, defaultCalling, ccList) {
  var p = dhbStr_(phoneRaw);
  if (!p) return "";

  // avoid "=010..." from CSV
  if (p.charAt(0) === "=") p = p.slice(1).trim();

  // remove quotes
  p = p.replace(/^"+|"+$/g, "");

  p = dhbToLatinDigits_(p);
  p = p.replace(/[()\-\s]/g, "");

  if (p.indexOf("00") === 0) p = "+" + p.slice(2);

  if (p.charAt(0) === "+") {
    var only = "+" + p.slice(1).replace(/\D/g, "");
    return only.length > 4 ? only : "";
  }

  var digits = p.replace(/\D/g, "");
  if (!digits) return "";

  // if digits already start with a known calling code, accept as +digits
  if (ccList && ccList.length) {
    for (var i = 0; i < ccList.length; i++) {
      if (digits.indexOf(String(ccList[i].calling)) === 0) return "+" + digits;
    }
  }

  // egypt local heuristic
  var d0 = digits;
  var looksEgyptLocal =
    (d0.indexOf("01") === 0 && d0.length === 11) ||
    (d0.indexOf("1") === 0 && d0.length === 10);

  if (looksEgyptLocal) {
    var dc = String(defaultCalling || "20").replace("+", "").trim();
    var local = d0.replace(/^0+/, "");
    var e164 = "+" + dc + local;
    return e164.length > 4 ? e164 : "";
  }

  if (digits.length >= 8) return "+" + digits;
  return "";
}

function dhbLookupCountry_(phoneE164, ccList) {
  if (!phoneE164) return { iso2: "", name_ar: "", name_en: "", flag: "🌍" };
  var digits = String(phoneE164).replace(/\D/g, "");
  if (!digits) return { iso2: "", name_ar: "", name_en: "", flag: "🌍" };

  for (var i = 0; i < ccList.length; i++) {
    var c = ccList[i];
    if (digits.indexOf(String(c.calling)) === 0) {
      return { iso2: c.iso2, name_ar: c.name_ar, name_en: c.name_en, flag: c.flag || "🌍" };
    }
  }
  return { iso2: "", name_ar: "", name_en: "", flag: "🌍" };
}

/* =========================
 * CORE: Misc
 * ========================= */

function dhbIsCancelled_(status) {
  return String(status || "").toLowerCase().indexOf("cancel") >= 0;
}

/**
 * Extract unit code from listing string (best-effort).
 * Example: "DHB - 101 | Downtown" => "101"
 */
function dhbExtractUnitCode_(listingName) {
  var s = dhbStr_(listingName);
  if (!s) return "";
  var m = s.match(/-\s*([A-Za-z0-9]+)\s*(\||$)/);
  if (m && m[1]) return String(m[1]).trim();
  var parts = s.split(/\s+/);
  return parts.length ? parts[parts.length - 1] : "";
}

/* =========================
 * SELF TEST (run once)
 * ========================= */

function DHB_CORE_SELF_TEST() {
  var ss = dhbSS_();
  Logger.log("Spreadsheet: " + ss.getName());

  // country codes load test (won't fail if sheet not created yet)
  var list = dhbLoadCountryCodes_();
  Logger.log("CountryCodes count: " + list.length);

  // date parse test
  var d = dhbParseDateLoose_("2026-01-19");
  Logger.log("Parsed: " + dhbToISODate_(d));

  return { ok: true, countryCodes: list.length };
}
