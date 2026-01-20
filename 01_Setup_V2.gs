/**
 * File: 01_Setup_V2.gs
 * DHB STAYS PMS v2 — Premium Spreadsheet Setup
 *
 * Depends on: 00_Core.gs
 *
 * Run:
 *  - DHB_SETUP_V2();                 // safe (create/format only)
 *  - DHB_SETUP_V2("soft_reset");     // clears operational data only
 *  - DHB_SETUP_V2("hard_reset");     // clears almost everything (keeps Users/Settings/CountryCodes/README)
 */

function DHB_SETUP_V2(mode) {
  var ss = dhbSS_();

  var resetMode = String(mode || "safe").toLowerCase(); // safe | soft_reset | hard_reset
  var DO_SOFT = (resetMode === "soft_reset");
  var DO_HARD = (resetMode === "hard_reset");

  // Spreadsheet settings
  try { ss.setSpreadsheetTimeZone(DHB.TZ); } catch (e) {}
  try { ss.setSpreadsheetLocale(DHB.LOCALE); } catch (e) {}

  // Theme
  var THEME = {
    font: "Arial",
    headerBg: "#111827",
    headerFg: "#FFFFFF",
    border: "#E5E7EB",
    gridHidden: true,
    bandHeader: "#F3F4F6",
    bandOdd: "#FAFAFA",
    bandEven: "#FFFFFF",
    shadowTab: {
      README: "#0EA5E9",
      Settings: "#6366F1",
      Users: "#8B5CF6",
      Units: "#22C55E",
      Reservations: "#EF4444",
      Guests: "#F59E0B",
      Prices: "#14B8A6",
      MaintenanceLog: "#F97316",
      Listings: "#A78BFA",
      CountryCodes: "#06B6D4",
      CleaningRequests: "#0F766E",
      UnitDailyChecklist: "#1D4ED8",
      UnitStatusHistory: "#334155",
      Imports_Raw: "#94A3B8",
      Import_Errors: "#FB7185",
      AuditLog: "#64748B"
    }
  };

  // Specs (FINAL v2)
  var SPECS = [
    specREADME_(THEME),
    specSettings_(THEME),
    specUsers_(THEME),
    specUnits_(THEME),
    specReservations_(THEME),
    specGuests_(THEME),
    specPrices_(THEME),
    specMaintenance_(THEME),
    specListings_(THEME),
    specCountryCodes_(THEME),
    specCleaningRequests_(THEME),
    specChecklist_(THEME),
    specStatusHistory_(THEME),
    specImportsRaw_(THEME),
    specImportErrors_(THEME),
    specAuditLog_(THEME)
  ];

  // Create/Format each sheet
  for (var i = 0; i < SPECS.length; i++) {
    var spec = SPECS[i];
    var sh = dhbEnsureSheet_(spec.name);

    // tab color
    try { sh.setTabColor(spec.tabColor); } catch (e1) {}

    // move into order (best effort)
    try { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); } catch (e2) {}

    // base style
    setupBaseStyle_(sh, THEME);

    // apply headers
    applyHeaders_(sh, spec.headers, THEME);

    // widths
    applyWidths_(sh, spec.widths);

    // freeze
    try { sh.setFrozenRows(1); } catch (e3) {}

    // hide gridlines
    try { sh.setHiddenGridlines(THEME.gridHidden); } catch (e4) {}

    // banding
    applyBanding_(sh, spec.headers.length, THEME);

    // filter row (skip README)
    if (spec.name !== DHB.SHEETS.README) {
      try {
        var f = sh.getFilter();
        if (f) f.remove();
        sh.getRange(1, 1, 1, spec.headers.length).createFilter();
      } catch (e5) {}
    }

    // formats
    if (spec.formatsFn) {
      try { spec.formatsFn(sh); } catch (e6) {}
    }

    // validations
    if (spec.validationsFn) {
      try { spec.validationsFn(sh); } catch (e7) {}
    }

    // protect header (warning only)
    protectHeader_(sh);

    // reset behavior
    if (DO_SOFT || DO_HARD) {
      if (DO_SOFT && shouldClearSoft_(spec.name)) {
        clearBelowHeader_(sh);
      }
      if (DO_HARD && shouldClearHard_(spec.name)) {
        clearBelowHeader_(sh);
      }
      // re-apply headers after clear
      applyHeaders_(sh, spec.headers, THEME);
    }

    // seed (only if empty OR hard reset OR spec says force on reset)
    if (spec.seedFn) {
      try { spec.seedFn(sh, DO_HARD); } catch (e8) {}
    }
  }

  // Conditional formatting (best effort)
  try { applyConditionalFormattingV2_(ss); } catch (e9) {}

  // Named ranges (best effort)
  try { applyNamedRangesV2_(ss); } catch (e10) {}

  SpreadsheetApp.flush();

  return { ok: true, mode: resetMode };
}

/* =========================
 * Reset rules
 * ========================= */

function shouldClearSoft_(name) {
  // soft_reset: operational only
  var list = [
    DHB.SHEETS.RES,
    DHB.SHEETS.GUESTS,
    DHB.SHEETS.CLEAN_REQ,
    DHB.SHEETS.CHECKLIST,
    DHB.SHEETS.RAW,
    DHB.SHEETS.ERRORS,
    DHB.SHEETS.AUDIT
  ];
  return list.indexOf(name) >= 0;
}

function shouldClearHard_(name) {
  // hard_reset: keep only README/Settings/Users/CountryCodes
  var keep = [DHB.SHEETS.README, DHB.SHEETS.SETTINGS, DHB.SHEETS.USERS, DHB.SHEETS.COUNTRY];
  return keep.indexOf(name) < 0;
}

/* =========================
 * Sheet Specs
 * ========================= */

function specREADME_(THEME) {
  return {
    name: DHB.SHEETS.README,
    tabColor: THEME.shadowTab.README,
    headers: ["DHB STAYS PMS (v2)", "Notes"],
    widths: [280, 980],
    seedFn: seedReadmeV2_
  };
}

function specSettings_(THEME) {
  return {
    name: DHB.SHEETS.SETTINGS,
    tabColor: THEME.shadowTab.Settings,
    headers: ["key", "value", "notes"],
    widths: [260, 260, 880],
    seedFn: seedSettingsV2_
  };
}

function specUsers_(THEME) {
  return {
    name: DHB.SHEETS.USERS,
    tabColor: THEME.shadowTab.Users,
    headers: ["username", "pin", "role", "is_active", "last_login", "notes"],
    widths: [220, 120, 180, 120, 180, 520],
    validationsFn: function (sh) {
      setValidationList_(sh, 2, DHB.ROLES);
      setValidationList_(sh, 3, ["TRUE", "FALSE"]);
    }
  };
}

function specUnits_(THEME) {
  return {
    name: DHB.SHEETS.UNITS,
    tabColor: THEME.shadowTab.Units,
    headers: [
      "unit_id","unit_code","unit_name",
      "status","cleaning_state","cleaning_notes",
      "color_hex","type","building","floor",
      "max_adults","max_children","beds","bathrooms","notes"
    ],
    widths: [90,110,240,130,140,260,110,110,140,80,110,120,80,100,360],
    validationsFn: function (sh) {
      setValidationList_(sh, 3, ["Active","Maintenance","Inactive"]);
      setValidationList_(sh, 4, ["Dirty","InProgress","Clean","Inspected","Maintenance"]);
      setValidationHexColor_(sh, 6);
      setValidationList_(sh, 7, ["Studio","1BR","2BR","3BR","Villa","Other"]);
    }
  };
}

function specReservations_(THEME) {
  return {
    name: DHB.SHEETS.RES,
    tabColor: THEME.shadowTab.Reservations,
    headers: [
      "res_id","channel","status","unit_id","guest_id","guest_name",
      "phone_raw","phone_e164","country_iso2",
      "check_in","check_out","nights",
      "adults","children","infants","guests_total",
      "booked_at","earnings_usd",
      "created_at","updated_at","source","notes"
    ],
    widths: [
      150,110,120,90,120,220,
      160,170,120,
      120,120,90,
      90,90,90,110,
      140,130,
      170,170,160,360
    ],
    formatsFn: function (sh) {
      setColumnNumberFormat_(sh, 9,  "yyyy-mm-dd");
      setColumnNumberFormat_(sh, 10, "yyyy-mm-dd");
      setColumnNumberFormat_(sh, 11, "0");
      setColumnNumberFormat_(sh, 12, "0");
      setColumnNumberFormat_(sh, 13, "0");
      setColumnNumberFormat_(sh, 14, "0");
      setColumnNumberFormat_(sh, 15, "0");
      setColumnNumberFormat_(sh, 16, "yyyy-mm-dd");
      setColumnNumberFormat_(sh, 17, "$#,##0.00");
      setColumnNumberFormat_(sh, 18, "yyyy-mm-dd hh:mm");
      setColumnNumberFormat_(sh, 19, "yyyy-mm-dd hh:mm");
    },
    validationsFn: function (sh) {
      setValidationList_(sh, 1, ["Airbnb","Booking","Manual"]);
      setValidationList_(sh, 2, ["Booked","Cancelled","Completed"]);
    }
  };
}

function specGuests_(THEME) {
  return {
    name: DHB.SHEETS.GUESTS,
    tabColor: THEME.shadowTab.Guests,
    headers: [
      "guest_id","name","phone_raw","phone_e164",
      "country_iso2","country_name_ar","country_name_en","flag_emoji",
      "total_spent_usd","visits_count","last_stay_date",
      "stayed_units","notes"
    ],
    widths: [120,240,160,170,120,180,200,110,150,120,140,360,360],
    formatsFn: function (sh) {
      setColumnNumberFormat_(sh, 8, "$#,##0.00");
      setColumnNumberFormat_(sh, 9, "0");
      setColumnNumberFormat_(sh, 10, "yyyy-mm-dd");
    }
  };
}

function specPrices_(THEME) {
  return {
    name: DHB.SHEETS.PRICES,
    tabColor: THEME.shadowTab.Prices,
    headers: ["price_id","unit_id","from_date","to_date","price_usd","note","created_by","created_at"],
    widths: [120,90,120,120,130,360,180,180],
    formatsFn: function (sh) {
      setColumnNumberFormat_(sh, 2, "yyyy-mm-dd");
      setColumnNumberFormat_(sh, 3, "yyyy-mm-dd");
      setColumnNumberFormat_(sh, 4, "$#,##0.00");
      setColumnNumberFormat_(sh, 7, "yyyy-mm-dd hh:mm");
    }
  };
}

function specMaintenance_(THEME) {
  return {
    name: DHB.SHEETS.MAINT,
    tabColor: THEME.shadowTab.MaintenanceLog,
    headers: ["maint_id","unit_id","from_date","to_date","reason","created_by","created_at","notes"],
    widths: [120,90,120,120,320,180,180,360],
    formatsFn: function (sh) {
      setColumnNumberFormat_(sh, 2, "yyyy-mm-dd");
      setColumnNumberFormat_(sh, 3, "yyyy-mm-dd");
      setColumnNumberFormat_(sh, 6, "yyyy-mm-dd hh:mm");
    }
  };
}

function specListings_(THEME) {
  return {
    name: DHB.SHEETS.LISTINGS,
    tabColor: THEME.shadowTab.Listings,
    headers: ["channel","listing_name","unit_id","notes"],
    widths: [120,520,100,420],
    validationsFn: function (sh) {
      setValidationList_(sh, 0, ["Airbnb","Booking","Manual"]);
    }
  };
}

function specCountryCodes_(THEME) {
  return {
    name: DHB.SHEETS.COUNTRY,
    tabColor: THEME.shadowTab.CountryCodes,
    headers: ["calling_code","iso2","name_ar","name_en","flag"],
    widths: [130,80,240,280,90],
    seedFn: seedCountryCodes80V2_
  };
}

function specCleaningRequests_(THEME) {
  return {
    name: DHB.SHEETS.CLEAN_REQ,
    tabColor: THEME.shadowTab.CleaningRequests,
    headers: [
      "id","unit_id","request_by","request_role","request_time",
      "severity","message","status",
      "seen_by","seen_at",
      "completed_by","completed_at"
    ],
    widths: [120,90,180,140,180,120,520,140,180,180,180,180],
    formatsFn: function (sh) {
      setColumnNumberFormat_(sh, 4,  "yyyy-mm-dd hh:mm");
      setColumnNumberFormat_(sh, 9,  "yyyy-mm-dd hh:mm");
      setColumnNumberFormat_(sh, 11, "yyyy-mm-dd hh:mm");
    },
    validationsFn: function (sh) {
      setValidationList_(sh, 5, ["Low","Normal","High","Urgent"]);
      setValidationList_(sh, 7, ["Pending","Seen","Completed","Cancelled"]);
    }
  };
}

function specChecklist_(THEME) {
  return {
    name: DHB.SHEETS.CHECKLIST,
    tabColor: THEME.shadowTab.UnitDailyChecklist,
    headers: [
      "date","unit_id","res_id","guest_id","guest_name",
      "hk_ready","fo_received","sec_checkedin",
      "problem_flag","problem_code","problem_reason",
      "updated_by","updated_at"
    ],
    widths: [120,90,140,120,240,110,130,140,120,180,420,180,180],
    formatsFn: function (sh) {
      setColumnNumberFormat_(sh, 0,  "yyyy-mm-dd");
      setColumnNumberFormat_(sh, 12, "yyyy-mm-dd hh:mm");
    },
    validationsFn: function (sh) {
      // Use checkboxes for the 4 flags
      setValidationCheckbox_(sh, 5);
      setValidationCheckbox_(sh, 6);
      setValidationCheckbox_(sh, 7);
      setValidationCheckbox_(sh, 8);

      setValidationList_(sh, 9, [
        "", "Unit Not Ready","Cleaning Delay","Guest No-show","ID Missing",
        "Payment Issue","Overbooking / Conflict","Maintenance Issue","Other"
      ]);
    }
  };
}

function specStatusHistory_(THEME) {
  return {
    name: DHB.SHEETS.STATUS_HIST,
    tabColor: THEME.shadowTab.UnitStatusHistory,
    headers: ["timestamp","unit_id","field","old_value","new_value","actor","notes"],
    widths: [180,90,180,240,240,180,520],
    formatsFn: function (sh) {
      setColumnNumberFormat_(sh, 0, "yyyy-mm-dd hh:mm");
    }
  };
}

function specImportsRaw_(THEME) {
  return {
    name: DHB.SHEETS.RAW,
    tabColor: THEME.shadowTab.Imports_Raw,
    headers: ["(Paste CSV export here)","Notes"],
    widths: [360, 920],
    seedFn: seedImportsRawHintV2_
  };
}

function specImportErrors_(THEME) {
  return {
    name: DHB.SHEETS.ERRORS,
    tabColor: THEME.shadowTab.Import_Errors,
    headers: ["imported_at","row_index","error_type","message","raw_json"],
    widths: [180,110,160,620,620],
    formatsFn: function (sh) {
      setColumnNumberFormat_(sh, 0, "yyyy-mm-dd hh:mm");
    }
  };
}

function specAuditLog_(THEME) {
  return {
    name: DHB.SHEETS.AUDIT,
    tabColor: THEME.shadowTab.AuditLog,
    headers: ["timestamp","username","action","details_json"],
    widths: [180,180,220,980],
    formatsFn: function (sh) {
      setColumnNumberFormat_(sh, 0, "yyyy-mm-dd hh:mm");
    }
  };
}

/* =========================
 * Seeds
 * ========================= */

function seedReadmeV2_(sh, force) {
  // seed only if empty unless force
  var has = sh.getLastRow() > 2;
  if (has && !force) return;

  clearBelowHeader_(sh);
  sh.getRange(1, 1, 1, 2).setValues([["DHB STAYS PMS (v2)", "Setup completed"]]);

  var lines = [
    ["Purpose", "Google Sheets DB + Apps Script Backend + Premium HTML UI"],
    ["Roles", "Admin / FrontOffice / Housekeeping / Security / Viewer"],
    ["Housekeeping", "Uses Units.cleaning_state + CleaningRequests + UnitDailyChecklist.hk_ready"],
    ["Security", "Sees arrivals + 3-checklist (HK/FO/Security) + Problem flag"],
    ["Do not", "Do not edit headers after setup"]
  ];
  sh.getRange(3, 1, lines.length, 2).setValues(lines);

  try { sh.setRowHeight(1, 36); } catch (e) {}
  try { sh.setRowHeights(3, lines.length, 28); } catch (e2) {}
}

function seedSettingsV2_(sh, force) {
  var has = (sh.getLastRow() > 1 && String(sh.getRange(2, 1).getValue() || "").trim() !== "");
  if (has && !force) return;

  clearBelowHeader_(sh);

  var rows = [
    ["app_version", "v2.0", "Do not change"],
    ["timezone", DHB.TZ, ""],
    ["default_language", "AR", "AR or EN"],
    ["default_currency", "USD", ""],
    ["default_country_calling_code", "20", "Used when phone has no + prefix"],
    ["sounds_on", "TRUE", "TRUE/FALSE"],
    ["haptics_on", "TRUE", "TRUE/FALSE"],
    ["ui_style", "StyleB_Premium", "Premium rounded UI"],
    ["default_price_usd", "0", "Fallback if no price range"],
    ["alerts_poll_sec", "8", "UI polling interval for alerts"]
  ];
  sh.getRange(2, 1, rows.length, 3).setValues(rows);
}

function seedImportsRawHintV2_(sh, force) {
  var has = (sh.getLastRow() > 1 && String(sh.getRange(2, 1).getValue() || "").trim() !== "");
  if (has && !force) return;

  clearBelowHeader_(sh);
  sh.getRange(2, 1, 1, 2).setValues([["Paste your CSV table here starting A1 (include header row).", "Tip: Prefer UI upload to avoid = in numbers"]]);
}

function seedCountryCodes80V2_(sh, force) {
  var has = (sh.getLastRow() > 1 && String(sh.getRange(2, 1).getValue() || "").trim() !== "");
  if (has && !force) return;

  clearBelowHeader_(sh);

  // Approved 80 base list (extend later manually)
  var data = [
    ["966","SA","المملكة العربية السعودية","Saudi Arabia","🇸🇦"],
    ["20","EG","مصر","Egypt","🇪🇬"],
    ["971","AE","الإمارات العربية المتحدة","United Arab Emirates","🇦🇪"],
    ["965","KW","الكويت","Kuwait","🇰🇼"],
    ["968","OM","عُمان","Oman","🇴🇲"],
    ["974","QA","قطر","Qatar","🇶🇦"],
    ["973","BH","البحرين","Bahrain","🇧🇭"],
    ["962","JO","الأردن","Jordan","🇯🇴"],
    ["964","IQ","العراق","Iraq","🇮🇶"],
    ["961","LB","لبنان","Lebanon","🇱🇧"],
    ["970","PS","فلسطين","Palestine","🇵🇸"],
    ["963","SY","سوريا","Syria","🇸🇾"],
    ["967","YE","اليمن","Yemen","🇾🇪"],
    ["212","MA","المغرب","Morocco","🇲🇦"],
    ["213","DZ","الجزائر","Algeria","🇩🇿"],
    ["216","TN","تونس","Tunisia","🇹🇳"],
    ["218","LY","ليبيا","Libya","🇱🇾"],
    ["249","SD","السودان","Sudan","🇸🇩"],
    ["252","SO","الصومال","Somalia","🇸🇴"],
    ["253","DJ","جيبوتي","Djibouti","🇩🇯"],
    ["222","MR","موريتانيا","Mauritania","🇲🇷"],
    ["269","KM","جزر القمر","Comoros","🇰🇲"],
    ["1","US","الولايات المتحدة","United States","🇺🇸"],
    ["44","GB","المملكة المتحدة","United Kingdom","🇬🇧"],
    ["86","CN","الصين","China","🇨🇳"],
    ["7","RU","روسيا","Russia","🇷🇺"],
    ["49","DE","ألمانيا","Germany","🇩🇪"],
    ["33","FR","فرنسا","France","🇫🇷"],
    ["90","TR","تركيا","Turkey","🇹🇷"],
    ["91","IN","الهند","India","🇮🇳"],
    ["55","BR","البرازيل","Brazil","🇧🇷"],
    ["81","JP","اليابان","Japan","🇯🇵"],
    ["39","IT","إيطاليا","Italy","🇮🇹"],
    ["34","ES","إسبانيا","Spain","🇪🇸"],
    ["1","CA","كندا","Canada","🇨🇦"],
    ["380","UA","أوكرانيا","Ukraine","🇺🇦"],
    ["48","PL","بولندا","Poland","🇵🇱"],
    ["40","RO","رومانيا","Romania","🇷🇴"],
    ["31","NL","هولندا","Netherlands","🇳🇱"],
    ["32","BE","بلجيكا","Belgium","🇧🇪"],
    ["30","GR","اليونان","Greece","🇬🇷"],
    ["46","SE","السويد","Sweden","🇸🇪"],
    ["41","CH","سويسرا","Switzerland","🇨🇭"],
    ["43","AT","النمسا","Austria","🇦🇹"],
    ["351","PT","البرتغال","Portugal","🇵🇹"],
    ["47","NO","النرويج","Norway","🇳🇴"],
    ["45","DK","الدنمارك","Denmark","🇩🇰"],
    ["358","FI","فنلندا","Finland","🇫🇮"],
    ["353","IE","أيرلندا","Ireland","🇮🇪"],
    ["420","CZ","التشيك","Czechia","🇨🇿"],
    ["36","HU","المجر","Hungary","🇭🇺"],
    ["62","ID","إندونيسيا","Indonesia","🇮🇩"],
    ["92","PK","باكستان","Pakistan","🇵🇰"],
    ["880","BD","بنغلاديش","Bangladesh","🇧🇩"],
    ["60","MY","ماليزيا","Malaysia","🇲🇾"],
    ["63","PH","الفلبين","Philippines","🇵🇭"],
    ["84","VN","فيتنام","Vietnam","🇻🇳"],
    ["66","TH","تايلاند","Thailand","🇹🇭"],
    ["82","KR","كوريا الجنوبية","South Korea","🇰🇷"],
    ["61","AU","أستراليا","Australia","🇦🇺"],
    ["64","NZ","نيوزيلندا","New Zealand","🇳🇿"],
    ["65","SG","سنغافورة","Singapore","🇸🇬"],
    ["93","AF","أفغانستان","Afghanistan","🇦🇫"],
    ["98","IR","إيران","Iran","🇮🇷"],
    ["94","LK","سريلانكا","Sri Lanka","🇱🇰"],
    ["27","ZA","جنوب أفريقيا","South Africa","🇿🇦"],
    ["234","NG","نيجيريا","Nigeria","🇳🇬"],
    ["254","KE","كينيا","Kenya","🇰🇪"],
    ["251","ET","إثيوبيا","Ethiopia","🇪🇹"],
    ["233","GH","غانا","Ghana","🇬🇭"],
    ["221","SN","السنغال","Senegal","🇸🇳"],
    ["52","MX","المكسيك","Mexico","🇲🇽"],
    ["54","AR","الأرجنتين","Argentina","🇦🇷"],
    ["57","CO","كولومبيا","Colombia","🇨🇴"],
    ["56","CL","تشيلي","Chile","🇨🇱"],
    ["51","PE","بيرو","Peru","🇵🇪"],
    ["58","VE","فنزويلا","Venezuela","🇻🇪"],
    ["53","CU","كوبا","Cuba","🇨🇺"],
    ["507","PA","بنما","Panama","🇵🇦"],
    ["598","UY","أوروغواي","Uruguay","🇺🇾"]
  ];

  sh.getRange(2, 1, data.length, 5).setValues(data);
}

/* =========================
 * Conditional Formatting
 * ========================= */

function applyConditionalFormattingV2_(ss) {
  // Units
  var shU = ss.getSheetByName(DHB.SHEETS.UNITS);
  if (shU) {
    var rulesU = shU.getConditionalFormatRules() || [];
    var maxR = Math.max(2000, shU.getMaxRows());
    var rangeStatus = shU.getRange(2, 4, maxR - 1, 1);
    var rangeClean = shU.getRange(2, 5, maxR - 1, 1);

    rulesU.push(makeTextRule_(rangeStatus, "Active", "#DCFCE7"));
    rulesU.push(makeTextRule_(rangeStatus, "Maintenance", "#FFEDD5"));
    rulesU.push(makeTextRule_(rangeStatus, "Inactive", "#F1F5F9"));

    rulesU.push(makeTextRule_(rangeClean, "Dirty", "#FEE2E2"));
    rulesU.push(makeTextRule_(rangeClean, "InProgress", "#FEF3C7"));
    rulesU.push(makeTextRule_(rangeClean, "Clean", "#DCFCE7"));
    rulesU.push(makeTextRule_(rangeClean, "Inspected", "#DBEAFE"));
    rulesU.push(makeTextRule_(rangeClean, "Maintenance", "#E2E8F0"));

    shU.setConditionalFormatRules(rulesU);
  }

  // Reservations status
  var shR = ss.getSheetByName(DHB.SHEETS.RES);
  if (shR) {
    var rulesR = shR.getConditionalFormatRules() || [];
    var maxRR = Math.max(4000, shR.getMaxRows());
    var rangeSt = shR.getRange(2, 3, maxRR - 1, 1);
    rulesR.push(makeTextRule_(rangeSt, "Booked", "#DCFCE7"));
    rulesR.push(makeTextRule_(rangeSt, "Cancelled", "#FEE2E2"));
    rulesR.push(makeTextRule_(rangeSt, "Completed", "#E2E8F0"));
    shR.setConditionalFormatRules(rulesR);
  }

  // CleaningRequests severity/status
  var shC = ss.getSheetByName(DHB.SHEETS.CLEAN_REQ);
  if (shC) {
    var rulesC = shC.getConditionalFormatRules() || [];
    var maxRC = Math.max(2000, shC.getMaxRows());
    var sev = shC.getRange(2, 6, maxRC - 1, 1);
    var st2 = shC.getRange(2, 8, maxRC - 1, 1);

    rulesC.push(makeTextRule_(sev, "Urgent", "#FEE2E2"));
    rulesC.push(makeTextRule_(sev, "High", "#FEF3C7"));
    rulesC.push(makeTextRule_(sev, "Normal", "#DBEAFE"));
    rulesC.push(makeTextRule_(sev, "Low", "#E2E8F0"));

    rulesC.push(makeTextRule_(st2, "Pending", "#FEF3C7"));
    rulesC.push(makeTextRule_(st2, "Seen", "#DBEAFE"));
    rulesC.push(makeTextRule_(st2, "Completed", "#DCFCE7"));
    rulesC.push(makeTextRule_(st2, "Cancelled", "#E2E8F0"));

    shC.setConditionalFormatRules(rulesC);
  }

  // Checklist: highlight problem rows, green checkboxes
  var shK = ss.getSheetByName(DHB.SHEETS.CHECKLIST);
  if (shK) {
    var rulesK = shK.getConditionalFormatRules() || [];
    var maxRK = Math.max(4000, shK.getMaxRows());

    // whole row range
    var fullRow = shK.getRange(2, 1, maxRK - 1, shK.getLastColumn());

    // problem_flag (col I = 9th col => letter I)
    rulesK.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$I2=TRUE')
        .setBackground("#FEE2E2")
        .setRanges([fullRow])
        .build()
    );

    // hk/fo/sec green when TRUE
    rulesK.push(makeBoolTrueRule_(shK.getRange(2, 6, maxRK - 1, 1), "#DCFCE7"));
    rulesK.push(makeBoolTrueRule_(shK.getRange(2, 7, maxRK - 1, 1), "#DCFCE7"));
    rulesK.push(makeBoolTrueRule_(shK.getRange(2, 8, maxRK - 1, 1), "#DCFCE7"));

    shK.setConditionalFormatRules(rulesK);
  }
}

function makeTextRule_(range, text, bg) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(String(text))
    .setBackground(String(bg))
    .setRanges([range])
    .build();
}

function makeBoolTrueRule_(range, bg) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=TRUE') // works for checkbox TRUE in same cell range
    .setBackground(String(bg))
    .setRanges([range])
    .build();
}

/* =========================
 * Named Ranges
 * ========================= */

function applyNamedRangesV2_(ss) {
  setNamedRange_(ss, "NR_SETTINGS", DHB.SHEETS.SETTINGS);
  setNamedRange_(ss, "NR_USERS", DHB.SHEETS.USERS);
  setNamedRange_(ss, "NR_UNITS", DHB.SHEETS.UNITS);
  setNamedRange_(ss, "NR_RESERVATIONS", DHB.SHEETS.RES);
  setNamedRange_(ss, "NR_GUESTS", DHB.SHEETS.GUESTS);
  setNamedRange_(ss, "NR_COUNTRY_CODES", DHB.SHEETS.COUNTRY);
  setNamedRange_(ss, "NR_CLEANING_REQUESTS", DHB.SHEETS.CLEAN_REQ);
  setNamedRange_(ss, "NR_CHECKLIST", DHB.SHEETS.CHECKLIST);
  setNamedRange_(ss, "NR_AUDITLOG", DHB.SHEETS.AUDIT);
}

function setNamedRange_(ss, name, sheetName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return;

  var lc = sh.getLastColumn() || 1;
  var lr = sh.getLastRow() || 1;
  if (lr < 1) lr = 1;

  var range = sh.getRange(1, 1, lr, lc);

  // remove existing
  var nrs = ss.getNamedRanges();
  for (var i = 0; i < nrs.length; i++) {
    if (nrs[i].getName() === name) {
      try { nrs[i].remove(); } catch (e) {}
    }
  }
  ss.setNamedRange(name, range);
}

/* =========================
 * Styling helpers
 * ========================= */

function setupBaseStyle_(sh, THEME) {
  try {
    sh.getRange(1, 1, 200, Math.min(40, sh.getMaxColumns())).setFontFamily(THEME.font);
  } catch (e) {}
  try { sh.setRowHeight(1, 34); } catch (e2) {}
  try { sh.setRowHeights(2, 200, 24); } catch (e3) {}
}

function applyHeaders_(sh, headers, THEME) {
  var range = sh.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range
    .setFontWeight("bold")
    .setFontColor(THEME.headerFg)
    .setBackground(THEME.headerBg)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  try {
    range.setBorder(true, true, true, true, true, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
  } catch (e) {}
}

function applyWidths_(sh, widths) {
  for (var i = 0; i < widths.length; i++) {
    try { sh.setColumnWidth(i + 1, widths[i]); } catch (e) {}
  }
}

function applyBanding_(sh, colCount, THEME) {
  try {
    var full = sh.getRange(1, 1, sh.getMaxRows(), Math.max(colCount, 1));
    var bandings = full.getBandings();
    for (var i = 0; i < bandings.length; i++) {
      try { bandings[i].remove(); } catch (e) {}
    }
  } catch (e2) {}

  var rows = Math.max(300, sh.getMaxRows());
  var r = sh.getRange(1, 1, rows, colCount);
  var band = r.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  try { band.setHeaderRowColor(THEME.bandHeader); } catch (e3) {}
  try { band.setFirstRowColor(THEME.bandOdd); } catch (e4) {}
  try { band.setSecondRowColor(THEME.bandEven); } catch (e5) {}
}

function clearBelowHeader_(sh) {
  var lr = sh.getLastRow();
  var lc = sh.getLastColumn();
  if (lr <= 1 || lc <= 0) return;
  sh.getRange(2, 1, lr - 1, lc).clearContent();
}

function protectHeader_(sh) {
  try {
    var p = sh.getRange(1, 1, 1, sh.getLastColumn()).protect();
    p.setDescription("Protect header row");
    p.setWarningOnly(true);
  } catch (e) {}
}

/* =========================
 * Validations + formats
 * ========================= */

function setValidationList_(sh, colIndexZeroBased, allowed) {
  var col = colIndexZeroBased + 1;
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(allowed, true)
    .setAllowInvalid(true) // allow invalid to not break imports
    .build();
  sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(rule);
}

function setValidationCheckbox_(sh, colIndexZeroBased) {
  var col = colIndexZeroBased + 1;
  var rule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(rule);
}

function setValidationHexColor_(sh, colIndexZeroBased) {
  var col = colIndexZeroBased + 1;
  var letter = colToLetter_(col);
  var formula = '=OR(' + letter + '2="",REGEXMATCH(' + letter + '2,"^#([A-Fa-f0-9]{6})$"))';
  var rule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied(formula)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(rule);
}

function setColumnNumberFormat_(sh, colIndexZeroBased, fmt) {
  var col = colIndexZeroBased + 1;
  sh.getRange(2, col, sh.getMaxRows() - 1, 1).setNumberFormat(fmt);
}

function colToLetter_(col) {
  var temp = col, letter = "";
  while (temp > 0) {
    var mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
}
