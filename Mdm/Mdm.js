/*************************************************
 * MDM SYNC
 *
 * Pushes this event, and this event's vendor rows,
 * to the central Master Data Management (MDM)
 * spreadsheet.
 *************************************************/

/*************************************************
 * PUSH EVENT TO MDM
 *************************************************/
function pushEventToMDM() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const sheetId = sheet.getSheetId();

  const mdm = SpreadsheetApp.openById(
    CONFIG.MDM_SPREADSHEET_ID
  );

  const eventsSheet = mdm.getSheetByName(
    CONFIG.MDM_EVENTS_SHEET
  );

  const data = eventsSheet
    .getDataRange()
    .getValues();

  // Check if this Event Sheet is already registered
  for (let i = 1; i < data.length; i++) {

    if (String(data[i][1]) === String(sheetId)) {

      SpreadsheetApp
        .getUi()
        .alert("This event is already registered in MDM.");

      return;
    }

  }

  eventsSheet.appendRow([

    ss.getId(),
    sheet.getSheetId(),
    "",
    sheet.getRange(EVENT.NAME).getValue(),
    sheet.getRange(EVENT.DATE).getValue(),
    sheet.getRange(EVENT.COUNTRY).getValue(),
    sheet.getRange(EVENT.CITY).getValue()

  ]);

  SpreadsheetApp
    .getUi()
    .alert("Event pushed to MDM.");

}

/*************************************************
 * COLLECT EVENT VENDORS
 *
 * Reads every vendor row across all blocks on the
 * active event sheet into a flat array, ready to
 * queue up for MDM.
 *************************************************/
function collectEventVendors() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getActiveSheet();
  Logger.log("Sheet: " + eventSheet.getName());

  const spreadsheetId = ss.getId();

  const vendors = [];

  const blocks = findBlocks(eventSheet);
  Logger.log("Blocks found: " + blocks.length);
  Logger.log(JSON.stringify(blocks, null, 2));

  for (const block of blocks) {
    Logger.log("Block:");
    Logger.log(JSON.stringify(block, null, 2));

    const cols = getBlockColumns(eventSheet, block);

    const firstVendorRow = block.startRow + 3;
    const lastVendorRow = block.endRow - 1;

    for (let row = firstVendorRow; row <= lastVendorRow; row++) {

      const vendorId =
        eventSheet.getRange(row, cols.vendor_id).getValue();

      const vendorName =
        eventSheet.getRange(row, cols.vendor_name).getValue();

      Logger.log({
        row: row,
        vendorId: vendorId,
        vendorName: vendorName
      });

      if (!vendorId && !vendorName) {
        continue;
      }

      vendors.push({

        spreadsheetId,
        sheetId: eventSheet.getSheetId(),
        sheetName: eventSheet.getName(),
        row,

        vendorId,
        vendorName,

        vendorType:
          eventSheet.getRange(row, cols.vendor_type).getValue(),

        contactPerson:
          eventSheet.getRange(row, cols.contact_person).getValue(),

        email:
          eventSheet.getRange(row, cols.email).getValue(),

        website:
          eventSheet.getRange(row, cols.website).getValue(),

        portfolio:
          cols.portfolio
            ? eventSheet.getRange(row, cols.portfolio).getValue()
            : "",

        quotation:
          cols.quotation_raw
            ? eventSheet.getRange(row, cols.quotation_raw).getValue()
            : "",

        fullDayUsd:
          cols.full_day_usd
            ? eventSheet.getRange(row, cols.full_day_usd).getValue()
            : "",

        halfDayUsd:
          cols.half_day_usd
            ? eventSheet.getRange(row, cols.half_day_usd).getValue()
            : "",

        crewNumber:
          cols.crew_number
            ? eventSheet.getRange(row, cols.crew_number).getValue()
            : "",

        equipment:
          cols.equipment_list
            ? eventSheet.getRange(row, cols.equipment_list).getValue()
            : "",

        comments:
          cols.comments
            ? eventSheet.getRange(row, cols.comments).getValue()
            : "",

        decision:
          cols.decision
            ? eventSheet.getRange(row, cols.decision).getValue()
            : "",
        rating:
          cols.rating
            ? eventSheet.getRange(row, cols.rating).getValue()
            : "",

      });

    }

  }

  return vendors;

}

/*************************************************
 * SUBMIT VENDOR CHANGES TO MDM
 *************************************************/
function submitToMDM() {

  const vendors = collectEventVendors();

  const mdm = SpreadsheetApp.openById(
    CONFIG.MDM_SPREADSHEET_ID
  );

  const queue = mdm.getSheetByName(
    CONFIG.VENDOR_UPDATE_QUEUE
  );

  vendors.forEach(vendor => {

    queue.appendRow([
      vendor.spreadsheetId,
      vendor.sheetId,
      vendor.sheetName,
      vendor.row,
      vendor.vendorId,
      vendor.vendorName,
      vendor.vendorType
        ? vendor.vendorType.charAt(0).toUpperCase() +
          vendor.vendorType.slice(1)
        : "",
      vendor.contactPerson,
      vendor.email,
      vendor.website,
      vendor.portfolio || "",
      vendor.quotation || "",
      vendor.equipment || "",
      vendor.comments || "",
      vendor.decision || "",
      vendor.rating || "",
      "Pending",
      new Date()
    ]);

  });

}
