const SHEET_NAME = "responses";
const SECRET = "socar-elecle-ut-2026";

const HEADERS = [
  "submitted_at",
  "participant_id",
  "participant_age",
  "participant_gender",
  "participant_experience",
  "participant_frequency",
  "participant_phone",
  "phase1_a_battery_last_digit",
  "phase1_a_elapsed_seconds",
  "phase1_b_battery_last_digit",
  "phase1_b_elapsed_seconds",
  "phase2_battery",
  "phase2_elapsed_seconds",
  "phase3_feedback_button",
  "phase3_urgent_return_button",
  "phase3_help_button",
  "phase4_qr_choice",
  "phase4_qr_reason",
  "phase4_bike_choice",
  "phase4_bike_reason",
  "phase4_riding_choice",
  "phase4_riding_reason",
  "user_agent"
];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || "{}");

    if (payload.secret !== SECRET) {
      return jsonResponse({ ok: false, error: "unauthorized" });
    }

    const sheet = getOrCreateSheet();
    ensureHeaders(sheet);

    const row = HEADERS.map((key) => payload[key] ?? "");
    sheet.appendRow(row);

    return jsonResponse({ ok: true });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error) });
  }
}

function doGet() {
  return jsonResponse({ ok: true, message: "UT response collector is running." });
}

function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function ensureHeaders(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = firstRow.some((value) => value !== "");
  const headerMatches = HEADERS.every((header, index) => firstRow[index] === header);

  if (!hasHeaders || !headerMatches) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
