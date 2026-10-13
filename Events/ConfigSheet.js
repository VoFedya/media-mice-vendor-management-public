/*************************************************
 * CONFIG SHEET
 *
 * Writes the per-event row (event sheet ID + RFQ
 * form ID) into the "Config" sheet during event
 * creation. Read back later by FormSubmission and
 * the Emails template lookups.
 *************************************************/

/*************************************************
 * SAVE FORM CONFIG
 *************************************************/
function saveFormConfig(sheetId, formId) {

  const config = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Config");

  const events = config.getRange(2, 1, config.getLastRow() - 1, 1).getValues();

  for (let i = 0; i < events.length; i++) {
    if (!events[i][0]) {
      config.getRange(i + 2, 1, 1, 2).setValues([[
        sheetId,
        formId
      ]]);
      return;
    }
  }
}
