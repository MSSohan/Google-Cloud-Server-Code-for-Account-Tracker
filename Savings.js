function onEditSavings(e) {

  if (!e || !e.range) return;

  const sh = e.range.getSheet();

  if (sh.getName() !== "Transactions") return;

  const row = e.range.getRow();
  if (row < 2) return;

  const watchedColumns = [2, 3, 4, 5, 6, 7];
  if (!watchedColumns.includes(e.range.getColumn())) return;

  const ss = e.source;
  const sg = ss.getSheetByName("Savings Goals");

  if (!sg) {
    ss.toast("Savings Goals sheet not found!", "Error", 5);
    return;
  }

  const category    = String(sh.getRange(row, 4).getDisplayValue()).trim();
  const clr         = String(sh.getRange(row, 5).getDisplayValue()).trim();
  const description = String(sh.getRange(row, 3).getDisplayValue()).trim();

  const categoryStatusMap = {
    "[Savings]": "Savings",
    "[Profit]":  "Profit"
  };
  const eligibleCategories = Object.keys(categoryStatusMap);
  const isEligible = eligibleCategories.includes(category) && clr === "C";

  if (!description) return;

  const savingsRow = findSavingsRow(sg, description);

  // ----------------------------
  // REMOVE CASE
  // ----------------------------
  if (!isEligible) {
    if (savingsRow) {
      sg.getRange(savingsRow, 2).clearContent();
      sg.getRange(savingsRow, 3).clearContent();
      sg.getRange(savingsRow, 4).clearContent();
      sg.getRange(savingsRow, 7).clearContent();
      sg.getRange(savingsRow, 8).clearContent();

      ss.toast("Removed from Savings Goals", "Savings", 3);
    }
    return;
  }

  // ----------------------------
  // AUTO CREATE OR UPDATE
  // ----------------------------
  const targetRow = savingsRow || getNextSavingsRow(sg);

  sg.getRange(targetRow, 2).setValue(sh.getRange(row, 2).getValue()); // Date
  sg.getRange(targetRow, 3).setValue(categoryStatusMap[category]);    // Status
  sg.getRange(targetRow, 4).setValue(description);                    // Description
  sg.getRange(targetRow, 7).setValue(sh.getRange(row, 6).getValue()); // Payment
  sg.getRange(targetRow, 8).setValue(sh.getRange(row, 7).getValue()); // Deposit


  ss.toast("Savings Updated ✓", "Savings", 3);
}


// ----------------------------
// Find existing row in SG by Description (col 4)
// ----------------------------
function findSavingsRow(sheet, description) {

  const values = sheet.getRange("D2:D").getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0].trim() === description) {
      return i + 2;
    }
  }

  return null;
}


// ----------------------------
// Find next empty row
// ----------------------------
function getNextSavingsRow(sheet) {

  const values = sheet.getRange("D2:D").getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      return i + 3;
    }
  }

  return 2;
}
