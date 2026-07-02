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
  "phase1_order",
  "phase1_first_variant",
  "phase1_second_variant",
  "phase1_assignment_source",
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
  "prize_result",
  "user_agent"
];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || "{}");
    return jsonResponse(savePayload(payload));
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error) });
  }
}

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  if (params.action === "phase1_assignment") {
    const sheet = getOrCreateSheet();
    ensureHeaders(sheet);

    const participantRow = findParticipantRow(sheet, params.participant_id);
    if (participantRow) {
      const existingOrder = sheet.getRange(participantRow, headerIndex("phase1_order") + 1).getValue();
      if (existingOrder === "A_FIRST" || existingOrder === "B_FIRST") {
        const existingCounts = countPhase1Orders(sheet);
        return scriptResponse(params.callback, {
          ok: true,
          phase1_order: existingOrder,
          a_first_count: existingCounts.aFirst,
          b_first_count: existingCounts.bFirst
        });
      }
    }

    const counts = countPhase1Orders(sheet);
    const order = counts.aFirst <= counts.bFirst ? "A_FIRST" : "B_FIRST";
    reservePhase1Order(sheet, params.participant_id, order);

    return scriptResponse(params.callback, {
      ok: true,
      phase1_order: order,
      a_first_count: counts.aFirst,
      b_first_count: counts.bFirst
    });
  }

  if (params.action === "submit") {
    try {
      const payload = JSON.parse(params.payload || "{}");
      return scriptResponse(params.callback, savePayload(payload));
    } catch (error) {
      return scriptResponse(params.callback, { ok: false, error: String(error) });
    }
  }

  return jsonResponse({ ok: true, message: "UT response collector is running." });
}

function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function savePayload(payload) {
  if (payload.secret !== SECRET) {
    return { ok: false, error: "unauthorized" };
  }

  const sheet = getOrCreateSheet();
  ensureHeaders(sheet);

  const row = HEADERS.map((key) => payload[key] ?? "");
  const participantRow = findParticipantRow(sheet, payload.participant_id);
  if (participantRow) {
    sheet.getRange(participantRow, 1, 1, HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  return { ok: true };
}

function headerIndex(headerName) {
  return HEADERS.indexOf(headerName);
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

function countPhase1Orders(sheet) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const orderIndex = headerRow.indexOf("phase1_order");
  const counts = { aFirst: 0, bFirst: 0 };

  if (orderIndex < 0 || sheet.getLastRow() < 2) {
    return counts;
  }

  const values = sheet.getRange(2, orderIndex + 1, sheet.getLastRow() - 1, 1).getValues();
  values.forEach((row) => {
    if (row[0] === "A_FIRST") counts.aFirst += 1;
    if (row[0] === "B_FIRST") counts.bFirst += 1;
  });

  return counts;
}

function reservePhase1Order(sheet, participantId, order) {
  if (!participantId) return;

  const existingRow = findParticipantRow(sheet, participantId);
  const row = HEADERS.map(() => "");
  row[headerIndex("participant_id")] = participantId;
  row[headerIndex("phase1_order")] = order;
  row[headerIndex("phase1_first_variant")] = order === "B_FIRST" ? "B" : "A";
  row[headerIndex("phase1_second_variant")] = order === "B_FIRST" ? "A" : "B";
  row[headerIndex("phase1_assignment_source")] = "sheet_reserved";

  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function findParticipantRow(sheet, participantId) {
  if (!participantId || sheet.getLastRow() < 2) return null;

  const participantIndex = headerIndex("participant_id");
  const values = sheet.getRange(2, participantIndex + 1, sheet.getLastRow() - 1, 1).getValues();
  for (let index = 0; index < values.length; index += 1) {
    if (values[index][0] === participantId) {
      return index + 2;
    }
  }

  return null;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function scriptResponse(callback, data) {
  if (!callback || !/^[A-Za-z_$][\w$]*(\.[A-Za-z_$][\w$]*)*$/.test(callback)) {
    return jsonResponse(data);
  }

  return ContentService
    .createTextOutput(`${callback}(${JSON.stringify(data)});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
