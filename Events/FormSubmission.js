/*************************************************
 * FORM SUBMISSION
 *
 * Handles the RFQ Google Form's on-submit trigger:
 * resolves which event sheet the response belongs
 * to, then hands the data to VendorManager.
 *************************************************/

/*************************************************
 * FORM SUBMIT TRIGGER
 *************************************************/
function onFormSubmit(e) {
  processFormSubmission(e);
}

/*************************************************
 * PROCESS FORM SUBMISSION
 *************************************************/
function processFormSubmission(e) {

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {

    const responseSheet = e.range.getSheet().getName();

    const eventSheetId = getEventName(responseSheet);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const eventSheet =
      ss.getSheets().find(
        sheet => sheet.getSheetId() === Number(eventSheetId)
      );

    if (!eventSheet) {
      throw new Error(`Event sheet "${eventSheetId}" not found.`);
    }

    const f = CONFIG.FORM_FIELDS;
    Logger.log(JSON.stringify(e.namedValues, null, 2));
    const quotationField = "Quotation (Google Drive URL)";

    const data = {
      company: e.namedValues[f.COMPANY][0],
      type: e.namedValues[f.TYPE][0],
      email: e.namedValues[f.EMAIL][0],
      website: e.namedValues[f.WEBSITE][0],
      contact: e.namedValues[f.CONTACT][0],
      sample: e.namedValues[f.SAMPLE][0],
      quotation: e.namedValues[quotationField]?.filter(Boolean)[0] || "",
      fullDay: "",
      halfDay: "",
      crewNumber: "",
      equipment: e.namedValues[f.EQUIPMENT][0],
      comments: e.namedValues[f.COMMENTS][0]
    };

    processVendorSubmission(
      eventSheet,
      data
    );

  } finally {

    lock.releaseLock();

  }

}

/*************************************************
 * GET EVENT NAME
 *************************************************/
function getEventName(responseSheetName) {

  const config = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.CONFIG_SHEET);

  const lastRow = config.getLastRow();

  for (let row = 2; row <= lastRow; row++) {

    if (
      String(config.getRange(row, 3).getValue()).trim() ===
      String(responseSheetName).trim()
    ) {
      return config.getRange(row, 1).getValue();
    }

  }

  return null;

}