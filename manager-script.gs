// manager-script.gs — Manager feedback helpers.
// Lives in the SAME Apps Script project as Code.gs, which owns doGet/doPost and
// routes here: GET ?action=getAll → managerList_(), POST {action:'managerSubmit'} → managerCreate_().
// This file defines NO doGet/doPost/json_ (Code.gs owns those) to avoid overriding them.
const MSHEET = 'Managers-Submissions';

function managerSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MSHEET);
  if (!sh) {
    sh = ss.insertSheet(MSHEET);
    sh.appendRow(['Timestamp','Department','Manager','Date','Lang','#Employees','JSON']);
  }
  return sh;
}

// READ: full JSON blob per submission (col 7), for round-trip to the dashboard.
function managerList_() {
  const rows = managerSheet_().getDataRange().getValues().slice(1);
  return rows.map(r => JSON.parse(r[6])).filter(Boolean);
}

// DELETE ONE: id = 0-based getAll position → sheet row id + 2 (header + 1-based).
function managerDelete_(d) {
  managerSheet_().deleteRow(d.id + 2);
  return { ok: true };
}

// DELETE ALL (one request): clear every data row below the header.
function clearManager_() {
  const sheet = managerSheet_();
  const n = sheet.getLastRow() - 1;
  if (n > 0) sheet.deleteRows(2, n);
  return { ok: true };
}

// CREATE: one row — readable summary cols + the full JSON kept for round-trip.
function managerCreate_(d) {
  managerSheet_().appendRow([
    d.timestamp, d.department, d.manager, d.date,
    d.lang, (d.employees || []).length,
    JSON.stringify(d)
  ]);
  return { ok: true };
}
