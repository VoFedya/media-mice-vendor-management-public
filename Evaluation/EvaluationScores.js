/*************************************************
 * EVALUATION SCORES
 *
 * Pulls scores back from the evaluation
 * spreadsheet (see EvaluationPush.js) into the
 * Rating column of the event sheet's videography
 * block.
 *************************************************/

/*************************************************
 * PULL EVALUATION SCORES
 *************************************************/
function pullEvaluationScores() {

  const eventSS =
    SpreadsheetApp.getActiveSpreadsheet();

  const eventSheet =
    eventSS.getActiveSheet();

  const evaluationFormula =
    eventSheet
      .getRange(EVENT.EVALUATION_URL)
      .getFormula();

  const evaluationId =
    evaluationFormula
      .match(/[-\w]{25,}/)?.[0];

  if (!evaluationId) {
    throw new Error(
      "Evaluation spreadsheet not found."
    );
  }

  const evaluationSS =
    SpreadsheetApp.openById(
      evaluationId
    );

  const vendorSheets =
    evaluationSS
      .getSheets()
      .filter(
        sheet =>
          /^Vendor_\d+$/.test(
            sheet.getName()
          )
      );

  let updated = 0;

  const block =
    findBlocks(eventSheet).find(
      b => b.type === "videography"
    );

  if (!block) {
    throw new Error(
      "Videography block not found."
    );
  }

  const firstRow =
    block.startRow + 3;

  const lastRow =
    block.endRow - 1;

  const values =
    eventSheet
      .getRange(
        block.startRow + 2,
        1,
        block.endRow - block.startRow - 2,
        eventSheet.getLastColumn()
      )
      .getValues();

  const headers =
    values.shift();

  const vendorNameCol =
    headers.indexOf(
      FIELD_MAP.NAME.event
    );

  const ratingCol =
    headers.indexOf("Rating");

  if (
    vendorNameCol === -1 ||
    ratingCol === -1
  ) {
    throw new Error(
      "Vendor Name or Rating column not found."
    );
  }

  vendorSheets.forEach(vendorSheet => {

    const vendorName =
      String(
        vendorSheet
          .getRange("B2")
          .getValue()
      ).trim();

    if (!vendorName) {
      return;
    }

    const videoScore =
      Number(
        vendorSheet
          .getRange("J9")
          .getValue()
      );

    const pmScore =
      Number(
        vendorSheet
          .getRange("J10")
          .getValue()
      );

    if (
      isNaN(videoScore) ||
      isNaN(pmScore)
    ) {
      return;
    }

    const average =
      (videoScore + pmScore) / 2;

    for (
      let row = firstRow;
      row <= lastRow;
      row++
    ) {

      const eventVendorName =
        String(
          eventSheet
            .getRange(
              row,
              vendorNameCol + 1
            )
            .getValue()
        ).trim();

      if (
        eventVendorName.toLowerCase() ===
        vendorName.toLowerCase()
      ) {

        eventSheet
          .getRange(
            row,
            ratingCol + 1
          )
          .setValue(average);

        updated++;

        break;
      }

    }

  });

  SpreadsheetApp
    .getUi()
    .alert(
      `${updated} vendor${updated === 1 ? "" : "s"} score${updated === 1 ? "" : "s"} pulled from Evaluation.`
    );

}

function pushPricingToMDM() {

  const eventSS =
    SpreadsheetApp.getActiveSpreadsheet();

  const eventSheet =
    eventSS.getActiveSheet();

  /*
   * Find Evaluation Sheet
   */

  const evaluationFormula =
    eventSheet
      .getRange(EVENT.EVALUATION_URL)
      .getFormula();

  const evaluationId =
    evaluationFormula
      .match(/[-\w]{25,}/)?.[0];

  if (!evaluationId) {
    throw new Error(
      "Evaluation spreadsheet not found."
    );
  }

  const evaluationSS =
    SpreadsheetApp.openById(
      evaluationId
    );

  /*
   * Open MDM
   */

  const mdmSS = SpreadsheetApp.openById(CONFIG.MDM_SPREADSHEET_ID);

  const pricingSheet =
    mdmSS.getSheetByName(
      "Vendor_Pricing"
    );

  if (!pricingSheet) {
    throw new Error(
      "Vendor_Pricing sheet not found."
    );
  }

  /*
   * Find Event ID in MDM
   */

  const mdmEventsSheet =
    mdmSS.getSheetByName(
      "Events"
    );

  if (!mdmEventsSheet) {
    throw new Error(
      "Events sheet not found."
    );
  }

  const spreadsheetId =
    eventSS.getId();

  const sheetId =
    eventSheet.getSheetId();

  const events =
    mdmEventsSheet
      .getDataRange()
      .getValues();

  const eventRow =
    events.find(row =>
      String(row[0]).trim() === spreadsheetId &&
      String(row[1]).trim() === String(sheetId)
    );

  if (!eventRow) {
    throw new Error(
      "Event not found in MDM Events."
    );
  }

  const eventId =
    String(eventRow[2]).trim();

  if (!eventId) {
    throw new Error(
      "Event ID is missing in MDM Events."
    );
  }

  /*
   * Find Evaluation Vendor Sheets
   */

  const vendorSheets =
    evaluationSS
      .getSheets()
      .filter(sheet =>
        /^Vendor_\d+$/.test(
          sheet.getName()
        )
      );

  /*
   * Existing MDM Pricing
   */

  const existing =
    pricingSheet
      .getDataRange()
      .getValues();

  let pushed = 0;

  /*
   * Read Pricing from Evaluation
   */

  vendorSheets.forEach(sheet => {

    const vendorId =
      String(
        sheet
          .getRange("B4")
          .getValue()
      ).trim();

    if (!vendorId) {
      SpreadsheetApp.getUi().alert(
        `Vendor ID is missing in ${sheet.getName()}.`
      );
      return;
    }

    const country =
      sheet
        .getRange("D4")
        .getValue();

    const city =
      sheet
        .getRange("E4")
        .getValue();

    for (
      let row = 14;
      row <= 20;
      row++
    ) {

      const quoteItem =
        sheet
          .getRange(row, 9)
          .getValue(); // I

      const price =
        sheet
          .getRange(row, 10)
          .getValue(); // J

      if (
        !quoteItem ||
        price === ""
      ) {
        continue;
      }

      const newRow = [
        vendorId,
        country,
        city,
        "Videography",
        quoteItem,
        price,
        eventId
      ];

      let found = false;

      /*
       * Check existing pricing
       */

      for (
        let i = 1;
        i < existing.length;
        i++
      ) {

        const oldRow =
          existing[i];

        const sameVendor =
          String(oldRow[0]).trim() ===
          String(newRow[0]).trim();

        const sameCountry =
          String(oldRow[1]).trim().toLowerCase() ===
          String(newRow[1]).trim().toLowerCase();

        const sameCity =
          String(oldRow[2]).trim().toLowerCase() ===
          String(newRow[2]).trim().toLowerCase();

        const sameCategory =
          String(oldRow[3]).trim().toLowerCase() ===
          String(newRow[3]).trim().toLowerCase();

        const sameItem =
          String(oldRow[4]).trim().toLowerCase() ===
          String(newRow[4]).trim().toLowerCase();

        if (
          sameVendor &&
          sameCountry &&
          sameCity &&
          sameCategory &&
          sameItem
        ) {

          pricingSheet
            .getRange(i + 1, 6)
            .setValue(newRow[5]);

          pricingSheet
            .getRange(i + 1, 7)
            .setValue(newRow[6]);

          found = true;
          pushed++;

          break;
        }

      }

      /*
       * Add new pricing record
       */

      if (!found) {

        pricingSheet
          .appendRow(newRow);

        existing.push(newRow);

        pushed++;

      }

    }

  });

  SpreadsheetApp
    .getUi()
    .alert(
      `${pushed} pricing record${pushed === 1 ? "" : "s"} pushed to MDM.`
    );

}