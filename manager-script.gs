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

// CREATE: one row — readable summary cols + the full JSON kept for round-trip.
function managerCreate_(d) {
  managerSheet_().appendRow([
    d.timestamp, d.department, d.manager, d.date,
    d.lang, (d.employees || []).length,
    JSON.stringify(d)
  ]);
  return { ok: true };
}
