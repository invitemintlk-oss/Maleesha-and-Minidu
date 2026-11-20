/**
 * Maleesha & Minidu — Wedding RSVP + Wishes (Google Apps Script Web App)
 *
 * Spreadsheet:
 * https://docs.google.com/spreadsheets/d/1WWCf8XcT8A-S3Vg8ZVFLPm1xm5W3xzmwk1teP4mqoNM/edit
 *
 * Two tabs are created automatically, each with formatted headers:
 *   • "RSVP"   → Timestamp | Name | Attendance | Guest Count | Has Wish
 *   • "Wishes" → Timestamp | Name | Attendance | Wish
 * Nothing needs to be set up in the spreadsheet by hand.
 *
 * Deploy:
 *   1. Open the spreadsheet → Extensions → Apps Script.
 *   2. Replace everything in Code.gs with this file and click Save.
 *   3. Run `setup` once from the editor and approve the permissions (creates both tabs).
 *   4. Deploy → New deployment → type "Web app"
 *        Execute as: Me
 *        Who has access: Anyone
 *   5. Copy the Web app URL (ends in /exec) into GOOGLE_SCRIPT_URL in src/App.tsx.
 *
 * Wishes are only stored in the sheet — they are never shown on the invitation website.
 */

const SPREADSHEET_ID = "1WWCf8XcT8A-S3Vg8ZVFLPm1xm5W3xzmwk1teP4mqoNM";

const SHEETS = {
  rsvp: {
    name: "RSVP",
    headers: ["Timestamp", "Name", "Attendance", "Guest Count", "Has Wish"],
    widths: [170, 240, 130, 110, 100],
    centered: ["C2:E"],
    wrapped: [],
  },
  wishes: {
    name: "Wishes",
    headers: ["Timestamp", "Name", "Attendance", "Wish"],
    widths: [170, 240, 130, 520],
    centered: ["C2:C"],
    wrapped: ["D2:D"],
  },
};

const HEADER_BG = "#6b1d3a";
const HEADER_FG = "#ffffff";
const MAX_GUESTS = 10;
const MAX_TEXT_LENGTH = 1000;

/* ------------------------------------------------------------------ */
/* Entry points                                                        */
/* ------------------------------------------------------------------ */

function doPost(e) {
  return handleRequest_(e);
}

function doGet(e) {
  // Opening the /exec URL in a browser creates both tabs and confirms the deployment works.
  if (!e || !e.parameter || !e.parameter.action) {
    setup();
    return jsonResponse_({ ok: true, message: "RSVP & Wishes service is running. Sheets are ready." });
  }
  return handleRequest_(e);
}

/** Run once from the editor to create the RSVP and Wishes tabs with headers. */
function setup() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  getSheet_(spreadsheet, SHEETS.rsvp);
  getSheet_(spreadsheet, SHEETS.wishes);
  removeEmptyDefaultSheet_(spreadsheet);
}

/* ------------------------------------------------------------------ */
/* Request handling                                                    */
/* ------------------------------------------------------------------ */

function handleRequest_(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const params = parseParams_(e);
    const action = String(params.action || "").trim().toLowerCase();

    if (action === "rsvp") return jsonResponse_(saveRsvp_(params));
    if (action === "wish") return jsonResponse_(saveWishOnly_(params));

    return jsonResponse_({ ok: false, message: "Invalid action" });
  } catch (error) {
    return jsonResponse_({
      ok: false,
      message: error && error.message ? error.message : "Unexpected error",
    });
  } finally {
    lock.releaseLock();
  }
}

/** Saves the RSVP row, and the wish (if any) to the separate Wishes tab. */
function saveRsvp_(params) {
  const name = cleanText_(params.name);
  if (!name) return { ok: false, message: "Name is required" };

  const attendanceRaw = String(params.attendance || "").trim().toLowerCase();
  const declined =
    attendanceRaw === "declined" || attendanceRaw === "decline" || String(params.guests) === "0";
  const attendance = declined ? "Declined" : "Attending";

  let guests = parseInt(params.guests, 10);
  if (declined) {
    guests = 0;
  } else if (!(guests >= 1)) {
    guests = 1;
  } else if (guests > MAX_GUESTS) {
    guests = MAX_GUESTS;
  }

  const wish = cleanText_(params.wish);
  const now = new Date();
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

  const rsvpSheet = getSheet_(spreadsheet, SHEETS.rsvp);
  rsvpSheet.appendRow([now, name, attendance, guests, wish ? "Yes" : "No"]);
  colorAttendance_(rsvpSheet, rsvpSheet.getLastRow(), 3, declined);

  if (wish) {
    const wishSheet = getSheet_(spreadsheet, SHEETS.wishes);
    wishSheet.appendRow([now, name, attendance, wish]);
    colorAttendance_(wishSheet, wishSheet.getLastRow(), 3, declined);
  }

  removeEmptyDefaultSheet_(spreadsheet);
  return { ok: true, message: "RSVP saved" };
}

/** Saves a wish on its own (action=wish), without an RSVP row. */
function saveWishOnly_(params) {
  const name = cleanText_(params.name);
  const wish = cleanText_(params.wish);
  if (!name) return { ok: false, message: "Name is required" };
  if (!wish) return { ok: false, message: "Wish is required" };

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const wishSheet = getSheet_(spreadsheet, SHEETS.wishes);
  wishSheet.appendRow([new Date(), name, "—", wish]);

  removeEmptyDefaultSheet_(spreadsheet);
  return { ok: true, message: "Wish saved" };
}

/* ------------------------------------------------------------------ */
/* Sheet helpers                                                       */
/* ------------------------------------------------------------------ */

/** Returns the tab described by `config`, creating it and/or its header row if missing. */
function getSheet_(spreadsheet, config) {
  let sheet = spreadsheet.getSheetByName(config.name);
  if (!sheet) sheet = spreadsheet.insertSheet(config.name);

  if (sheet.getLastRow() === 0) {
    sheet
      .getRange(1, 1, 1, config.headers.length)
      .setValues([config.headers])
      .setFontWeight("bold")
      .setFontColor(HEADER_FG)
      .setBackground(HEADER_BG)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sheet.setRowHeight(1, 32);
    sheet.setFrozenRows(1);

    config.widths.forEach(function (width, i) {
      sheet.setColumnWidth(i + 1, width);
    });

    sheet.getRange("A2:A").setNumberFormat("yyyy-mm-dd hh:mm");
    config.centered.forEach(function (a1) {
      sheet.getRange(a1).setHorizontalAlignment("center");
    });
    config.wrapped.forEach(function (a1) {
      sheet.getRange(a1).setWrap(true).setVerticalAlignment("top");
    });
  }

  return sheet;
}

/** Deletes the default "Sheet1" tab once our tabs exist, but only if it is empty. */
function removeEmptyDefaultSheet_(spreadsheet) {
  const defaultSheet = spreadsheet.getSheetByName("Sheet1");
  if (defaultSheet && defaultSheet.getLastRow() === 0 && spreadsheet.getSheets().length > 1) {
    spreadsheet.deleteSheet(defaultSheet);
  }
}

function colorAttendance_(sheet, row, column, declined) {
  sheet
    .getRange(row, column)
    .setFontColor(declined ? "#a3263f" : "#1e7a46")
    .setFontWeight("bold");
}

/* ------------------------------------------------------------------ */
/* Parsing / utils                                                     */
/* ------------------------------------------------------------------ */

/** Trims, caps length, and neutralises spreadsheet formula injection. */
function cleanText_(value) {
  let text = String(value == null ? "" : value).trim().slice(0, MAX_TEXT_LENGTH);
  if (/^[=+\-@]/.test(text)) text = "'" + text;
  return text;
}

function parseParams_(e) {
  const params = Object.assign({}, (e && e.parameter) || {});

  if (!(e && e.postData && e.postData.contents)) return params;

  const contents = String(e.postData.contents || "").trim();
  if (!contents) return params;

  if (contents.charAt(0) === "{") {
    return Object.assign(params, JSON.parse(contents));
  }

  return Object.assign(params, parseFormEncoded_(contents));
}

function parseFormEncoded_(raw) {
  const result = {};
  const pairs = String(raw || "").split("&");

  for (let i = 0; i < pairs.length; i++) {
    if (!pairs[i]) continue;
    const parts = pairs[i].split("=");
    const key = decodeURIComponent((parts[0] || "").replace(/\+/g, " "));
    const value = decodeURIComponent((parts.slice(1).join("=") || "").replace(/\+/g, " "));
    if (key) result[key] = value;
  }

  return result;
}

function jsonResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
