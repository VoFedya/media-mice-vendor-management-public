/*************************************************
 * EVALUATION PUSH
 *
 * Pushes selected (Decision = "Evaluate")
 * videography vendors from the event sheet into a
 * dedicated evaluation spreadsheet, one tab per
 * vendor, pre-filled with prior pricing from MDM.
 *************************************************/

/*************************************************
 * PUSH TO EVALUATION
 *************************************************/
function pushToEvaluation() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getActiveSheet();

  const vendors = getSelectedVideographyVendors();

  if (vendors.length === 0) {

    SpreadsheetApp.getUi().alert(
      "Please select at least one videography vendor."
    );

    return;

  }

  let evaluationSS;

  const evaluationUrl =
    sheet.getRange(EVENT.EVALUATION_URL).getFormula();

  if (evaluationUrl) {

    const evaluationId =
      evaluationUrl.match(/[-\w]{25,}/)[0];

    evaluationSS =
      SpreadsheetApp.openById(evaluationId);

  } else {

    evaluationSS =
      createEvaluationSpreadsheet();

  }

  const existingNames =
    evaluationSS.getSheets()
      .filter(sheet => /^Vendor_\d+$/.test(sheet.getName()))
      .map(sheet =>
        String(sheet.getRange("B2").getValue()).trim().toLowerCase()
      );

  const newVendors =
    vendors.filter(vendor =>
      !existingNames.includes(
        String(vendor.name).trim().toLowerCase()
      )
    );

  if (newVendors.length > 0) {

    createVendorSheets(
      evaluationSS,
      newVendors
    );

  }

  SpreadsheetApp.getUi().alert(
    newVendors.length === 0
      ? "No new vendors selected."
      : `${newVendors.length} new vendor${newVendors.length === 1 ? "" : "s"} added.`
  );

}

/*************************************************
 * GET SELECTED VIDEOGRAPHY VENDORS
 *************************************************/
function getSelectedVideographyVendors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const country = sheet.getRange("C2").getValue();
  const city = sheet.getRange("C3").getValue();

  const block = findBlocks(sheet).find(
    b => b.type === "videography"
  );

  if (!block) {
    throw new Error("Videography block not found.");
  }

  const values = sheet
    .getRange(
      block.startRow + 2,
      1,
      block.endRow - block.startRow - 2,
      sheet.getLastColumn()
    )
    .getValues();

  const headers = values.shift();

  const columns = {
    id: headers.indexOf(FIELD_MAP.ID.event),
    name: headers.indexOf(FIELD_MAP.NAME.event),
    website: headers.indexOf(FIELD_MAP.WEBSITE.event),
    contact: headers.indexOf(FIELD_MAP.CONTACT.event),
    portfolio: headers.indexOf(FIELD_MAP.PORTFOLIO.event),
    decision: headers.indexOf(FIELD_MAP.DECISION.event),
    email: headers.indexOf(FIELD_MAP.EMAIL.event),
    equipment: headers.indexOf(FIELD_MAP.Equipment.event),
    quotation: headers.indexOf(FIELD_MAP.QUOTATION.event),
  };

  const vendors = [];

  values.forEach(row => {
    if (row[columns.decision] !== "Evaluate") {
      return;
    }

    vendors.push({
      id: row[columns.id],
      name: row[columns.name],
      website: row[columns.website],
      contact: row[columns.contact],
      email: row[columns.email],
      portfolio: row[columns.portfolio],
      equipment: row[columns.equipment],
      quotation_raw: row[columns.quotation],
      country: country,
      city: city
    });
  });

  return vendors;
}

/*************************************************
 * CREATE EVALUATION SPREADSHEET
 *************************************************/
function createEvaluationSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const templateFile = DriveApp.getFileById(
    CONFIG.EVALUATION_TEMPLATE_ID
  );

  const folderCell =
    sheet.getRange(EVENT.FOLDER_URL);

  const folderUrl =
    folderCell
      .getFormula()
      .match(/HYPERLINK\("([^"]+)"/i)?.[1] || "";

  const folderId =
    folderUrl.match(/[-\w]{25,}/)?.[0];

  if (!folderId) {
    throw new Error("Event Folder URL not found.");
  }

  const eventFolder = DriveApp.getFolderById(folderId);

  const fileName = `${sheet.getName()} - Evaluation`;

  const copy = templateFile.makeCopy(fileName, eventFolder);

  const evaluationSS = SpreadsheetApp.openById(copy.getId());

  sheet
    .getRange(EVENT.EVALUATION_URL)
    .setFormula(
      `=HYPERLINK("${evaluationSS.getUrl()}","Evaluation")`
    );

  Logger.log(evaluationSS);
  return evaluationSS;
}

/*************************************************
 * CREATE VENDOR SHEETS
 *************************************************/
function createVendorSheets(ss, vendors) {

  if (!ss) {
    throw new Error("ss is undefined");
  }

  const template = ss.getSheetByName("Vendor_Template");

  if (!template) {
    throw new Error("Vendor_Template sheet not found.");
  }

  const templateIndex = template.getIndex();

  let i = 1;

  vendors.forEach(vendor => {

    while (ss.getSheetByName(`Vendor_${i}`)) {
      i++;
    }

    const sheet = template.copyTo(ss);
    sheet.showSheet();

    sheet.setName(`Vendor_${i}`);

    sheet.getRange("B2").setValue(vendor.name);
    sheet.getRange("B3").setValue(vendor.website);
    sheet.getRange("B4").setValue(vendor.id);
    sheet.getRange("D2").setValue(vendor.contact);
    sheet.getRange("E2").setValue(vendor.email);
    sheet.getRange("D3").setValue(vendor.portfolio);
    sheet.getRange("E24").setValue(vendor.equipment);
    sheet.getRange("I24").setValue(vendor.quotation_raw);
    sheet.getRange("D4").setValue(vendor.country);
    sheet.getRange("E4").setValue(vendor.city);

    const oldPricing = getOldVendorPricing(vendor);

    for (let row = 14; row <= 20; row++) {

      const quoteItem = String(
        sheet.getRange(row, 9).getValue() || ""
      ).trim();

      const price = oldPricing[quoteItem.toLowerCase()];

      if (price !== undefined) {
        sheet.getRange(row, 11).setValue(price);
      }

    }

    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(templateIndex + i);
  });

  template.hideSheet();
}

/*************************************************
 * GET OLD VENDOR PRICING
 *************************************************/
function getOldVendorPricing(vendor) {

  if (!vendor.id) {
    return {};
  }

  const mdmSS = SpreadsheetApp.openById(
    CONFIG.MDM_SPREADSHEET_ID
  );

  const pricingSheet = mdmSS.getSheetByName(
    "Vendor_Pricing"
  );

  if (!pricingSheet) {
    throw new Error("Vendor_Pricing sheet not found.");
  }

  const pricingData = pricingSheet
    .getDataRange()
    .getValues();

  const oldPricing = {};

  for (let i = 1; i < pricingData.length; i++) {

    const row = pricingData[i];

    const vendorId = String(row[0] || "").trim();
    const country = String(row[1] || "").trim();
    const city = String(row[2] || "").trim();
    const category = String(row[3] || "").trim();
    const quoteItem = String(row[4] || "").trim();
    const price = row[5];

    if (
      vendorId === String(vendor.id).trim() &&
      country.toLowerCase() ===
        String(vendor.country).trim().toLowerCase() &&
      city.toLowerCase() ===
        String(vendor.city).trim().toLowerCase() &&
      category.toLowerCase() === "videography"
    ) {

      oldPricing[quoteItem.toLowerCase()] = price;

    }

  }

  return oldPricing;
}
