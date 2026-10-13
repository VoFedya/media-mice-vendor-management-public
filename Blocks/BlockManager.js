/*************************************************
 * BLOCK MANAGER
 *
 * A "block" is a vendor-category section on an
 * event sheet (Videography, Printing, etc.),
 * delimited by CONFIG.BLOCK.START_MARKER /
 * END_MARKER rows. These functions find, create,
 * resize, and inspect blocks and their columns.
 *************************************************/

/*************************************************
 * APPEND BLOCK
 *************************************************/
function appendBlock(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const eventSheet = ss.getActiveSheet();
  const templateSheet = ss.getSheetByName(CONFIG.TEMPLATES[type]);

  if (!templateSheet) {
    throw new Error(`Template for "${type}" not found.`);
  }

  const rows = templateSheet.getLastRow();
  const cols = templateSheet.getLastColumn();

  const insertRow = eventSheet.getLastRow() + CONFIG.BLOCK.SPACING;

  eventSheet.insertRowsBefore(insertRow, rows);

  templateSheet
    .getRange(1, 1, rows, cols)
    .copyTo(eventSheet.getRange(insertRow, 1));
}


/*************************************************
 * FIND BLOCKS
 *************************************************/
function findBlocks(
  sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
) {
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();

  const blocks = [];

  for (let i = 0; i < values.length; i++) {
    const cell = String(values[i][0]).trim();

    if (cell !== CONFIG.BLOCK.START_MARKER) continue;

    const startRow = i + 1;

    const typeCell = String(values[i + 1]?.[0] || "").trim();

    if (!typeCell.startsWith("TYPE=")) {
      throw new Error(`Missing TYPE after block starting at row ${startRow}.`);
    }

    const type = typeCell.substring(5);

    let endRow = null;

    for (let j = i + 1; j < values.length; j++) {
      if (String(values[j][0]).trim() === CONFIG.BLOCK.END_MARKER) {
        endRow = j + 1;
        break;
      }
    }

    if (!endRow) {
      throw new Error(
        `Missing ${CONFIG.BLOCK.END_MARKER} for "${type}" block starting at row ${startRow}.`
      );
    }

    blocks.push({
      type,
      startRow,
      endRow
    });

    i = endRow - 1;
  }

  return blocks;
}


/*************************************************
 * GET CURRENT BLOCK
 *************************************************/
function getCurrentBlock() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const currentRow = sheet.getActiveCell().getRow();

  const blocks = findBlocks(sheet);

  for (const block of blocks) {
    if (currentRow >= block.startRow && currentRow <= block.endRow) {
      return block;
    }
  }

  return null;
}


/*************************************************
 * DELETE CURRENT BLOCK
 *************************************************/
function deleteCurrentBlock() {
  const block = getCurrentBlock();

  if (!block) {
    SpreadsheetApp.getUi().alert(
      "Please select a cell inside a block."
    );
    return;
  }

  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    "Delete Block",
    `Delete "${block.type}" block?\n\nThis action cannot be undone.`,
    ui.ButtonSet.OK_CANCEL
  );

  if (result !== ui.Button.OK) {
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const rowsInBlock = block.endRow - block.startRow + 1;
  let rowsToDelete = rowsInBlock;

  if (
    CONFIG.BLOCK.SPACING === 1 &&
    block.endRow < sheet.getMaxRows() &&
    sheet.getRange(block.endRow + 1, 1).isBlank()
  ) {
    rowsToDelete++;
  }

  sheet.deleteRows(block.startRow, rowsToDelete);
}


/*************************************************
 * ADD VENDOR ROWS
 *************************************************/
function addVendorRows() {
  const block = getCurrentBlock();

  if (!block) {
    SpreadsheetApp.getUi().alert(
      "Please select a cell inside a block."
    );
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  insertVendorRows(sheet, block);
}


/*************************************************
 * INSERT VENDOR ROWS
 *************************************************/
function insertVendorRows(sheet, block, rowsToAdd = 10) {
  const insertRow = block.endRow;

  sheet.insertRowsBefore(insertRow, rowsToAdd);

  const templateRow = insertRow - 1;
  const lastColumn = sheet.getLastColumn();

  const source = sheet.getRange(templateRow, 1, 1, lastColumn);
  const destination = sheet.getRange(insertRow, 1, rowsToAdd, lastColumn);

  source.copyTo(
    destination,
    SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
    false
  );

  destination.clearContent();

  return insertRow;
}


/*************************************************
 * GET DECISION COLUMN
 *************************************************/
function getDecisionColumn(block) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const values = sheet
    .getRange(
      block.startRow,
      1,
      block.endRow - block.startRow + 1,
      sheet.getLastColumn()
    )
    .getValues();

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (values[r][c] === "Decision") {
        return c + 1;
      }
    }
  }

  throw new Error("Decision column not found.");
}


/*************************************************
 * ADD PRICE COLUMN
 *************************************************/
function addPriceColumn() {
  const block = getCurrentBlock();

  if (!block) {
    SpreadsheetApp.getUi().alert(
      "Please select a cell inside a block."
    );
    return;
  }

  if (!CONFIG.DYNAMIC_PRICE_COLUMNS.includes(block.type)) {
    SpreadsheetApp.getUi().alert(
      `${block.type} doesn't support dynamic price columns.`
    );
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const decisionCol = getDecisionColumn(block);

  const numRows = block.endRow - block.startRow + 1;
  const width = sheet.getLastColumn() - decisionCol + 1;

  sheet
    .getRange(
      block.startRow,
      decisionCol,
      numRows,
      width
    )
    .moveTo(
      sheet.getRange(
        block.startRow,
        decisionCol + 1
      )
    );

  sheet
    .getRange(
      block.startRow,
      decisionCol - 1,
      numRows,
      1
    )
    .copyTo(
      sheet.getRange(
        block.startRow,
        decisionCol,
        numRows,
        1
      ),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );

  sheet
    .getRange(
      block.startRow,
      decisionCol,
      numRows,
      1
    )
    .clearDataValidations();

  sheet
    .getRange(
      block.startRow,
      decisionCol,
      numRows,
      1
    )
    .clearContent();
}


/*************************************************
 * GET VENDOR CATEGORIES
 *************************************************/
function getVendorCategories() {
  return Object.keys(CONFIG.TEMPLATES);
}


/*************************************************
 * GET BLOCK COLUMNS
 *
 * Reads the header row of a block and returns a
 * map of normalized header name -> column index.
 * Used by every feature that reads/writes vendor
 * rows (email sending, form submission, vendor
 * finder, MDM sync, etc.).
 *************************************************/
function getBlockColumns(sheet, block) {

  const headerRow = block.startRow + 2;

  const headers = sheet
    .getRange(headerRow, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const columns = {};

  headers.forEach((header, index) => {
    if (header) {
      columns[normalize(header)] = index + 1;
    }
  });

  return columns;
}
