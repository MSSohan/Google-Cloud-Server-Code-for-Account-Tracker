function onEditLiabilities(e) {

  if (!e || !e.range) return;

  const sh = e.range.getSheet();

  if (sh.getName() !== "Transactions") return;

  const row = e.range.getRow();
  if (row < 2) return;

  const watchedColumns = [2, 3, 4, 5, 6, 7];
  if (!watchedColumns.includes(e.range.getColumn())) return;

  const ss = e.source;
  const li = ss.getSheetByName("Liabilities");

  if (!li) {
    ss.toast("Liabilities sheet not found!", "Error", 5);
    return;
  }

  const category    = String(sh.getRange(row, 4).getDisplayValue()).trim();
  const clr         = String(sh.getRange(row, 5).getDisplayValue()).trim();
  const description = String(sh.getRange(row, 3).getDisplayValue()).trim();

  const categoryStatusMap = {
    "[Lent]":          "Lent",
    "[Received Back]": "Received Back",
    "[Borrowed]":      "Borrowed",
    "[Repaid]":        "Repaid"
  };

  const eligibleCategories = Object.keys(categoryStatusMap);
  const isEligible = eligibleCategories.includes(category) && clr === "C";

  if (!description) return;

  const liabRow = findLiabilitiesRow(li, description);

  // ----------------------------
  // REMOVE CASE
  // ----------------------------
  if (!isEligible) {
    if (liabRow) {
      li.getRange(liabRow, 2).clearContent();
      li.getRange(liabRow, 3).clearContent();
      li.getRange(liabRow, 4).clearContent();
      li.getRange(liabRow, 8).clearContent();
      li.getRange(liabRow, 9).clearContent();

      ss.toast("Removed from Liabilities", "Liabilities", 3);
    }
    return;
  }

  // ----------------------------
  // AUTO CREATE OR UPDATE
  // ----------------------------
  const targetRow = liabRow || getNextLiabilitiesRow(li);

  li.getRange(targetRow, 2).setValue(sh.getRange(row, 2).getValue()); // Date
  li.getRange(targetRow, 3).setValue(categoryStatusMap[category]);    // Status
  li.getRange(targetRow, 4).setValue(description);                    // Description
  li.getRange(targetRow, 8).setValue(sh.getRange(row, 6).getValue()); // Amount Out (Payment)
  li.getRange(targetRow, 9).setValue(sh.getRange(row, 7).getValue()); // Amount In (Deposit)

  ss.toast("Liabilities Updated ✓", "Liabilities", 3);
}


// ----------------------------
// Find existing row in Liabilities by Description (col 4)
// ----------------------------
function findLiabilitiesRow(sheet, description) {

  const values = sheet.getRange("D2:D").getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0].trim() === description) {
      return i + 2;
    }
  }

  return null;
}


// ----------------------------
// Find next empty row in Liabilities
// ----------------------------
function getNextLiabilitiesRow(sheet) {

  const values = sheet.getRange("D2:D").getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      return i + 3;
    }
  }

  return 2;
}
