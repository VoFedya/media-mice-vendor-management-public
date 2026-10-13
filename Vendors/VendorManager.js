/*************************************************
 * VENDOR MANAGER
 *
 * Reads and writes vendor rows inside a block,
 * and looks vendors up against the shared
 * Vendor_Cache sheet.
 *************************************************/

/*************************************************
 * PROCESS VENDOR SUBMISSION
 *************************************************/
function processVendorSubmission(
  eventSheet,
  data
) {

  const block = findBlocks(eventSheet).find(
    block => block.type.toLowerCase() === data.type.toLowerCase()
  );

  if (!block) {
    throw new Error(
      `Vendor block "${data.type}" not found.`
    );
  }

  let row = findVendorRow(
    eventSheet,
    block,
    data.company
  );

  if (row !== -1) {

    updateEventVendor(
      eventSheet,
      block,
      row,
      data
    );

    return;

  }

  row = findEmptyVendorRow(
    eventSheet,
    block
  );

  if (row === -1) {

    row = insertVendorRows(
      eventSheet,
      block,
      1
    );

  }

  addNewVendor(
    eventSheet,
    block,
    row,
    data
  );

}


/*************************************************
 * ADD NEW VENDOR
 *************************************************/
function addNewVendor(
  eventSheet,
  block,
  row,
  data
) {

  const cols = getBlockColumns(eventSheet, block);

  Logger.log(cols);
  Logger.log("quotation_raw col = " + cols["quotation_raw"]);
  Logger.log("quotation value = " + data.quotation);

  const vendor = findVendor(data.company);

  const values = {
    vendor_id: vendor ? vendor.id : "",
    vendor_name: data.company,
    vendor_status: vendor ? "Existing" : "New Vendor",
    vendor_type: data.type || (vendor ? vendor.type : ""),
    contact_person: data.contact || (vendor ? vendor.contact : ""),
    email: data.email || (vendor ? vendor.email : ""),
    website: data.website || (vendor ? vendor.website : ""),

    send_email: false,
    email_sent: "",
    last_activity: "",

    portfolio: data.sample,
    quotation_raw: data.quotation,
    full_day_usd: data.fullDay,
    half_day_usd: data.halfDay,
    crew_number: data.crewNumber,
    equipment_list: data.equipment,
    comments: data.comments
  };

  for (const [header, value] of Object.entries(values)) {

    const col = cols[header];

    if (col) {
      eventSheet.getRange(row, col).setValue(value);
    }

  }

  return row;

}


/*************************************************
 * UPDATE EVENT VENDOR
 *************************************************/
function updateEventVendor(
  eventSheet,
  block,
  row,
  data
) {

  const cols = getBlockColumns(eventSheet, block);

  const updates = {
    vendor_type: data.type,
    contact_person: data.contact,
    email: data.email,
    website: data.website,
    portfolio: data.sample,
    quotation_raw: data.quotation,
    full_day_usd: data.fullDay,
    half_day_usd: data.halfDay,
    crew_number: data.crewNumber,
    equipment_list: data.equipment,
    comments: data.comments
  };
  Logger.log(data.type);

  for (const [header, value] of Object.entries(updates)) {

    const col = cols[header];

    if (!col) {
      continue;
    }

    if (value !== "" && value !== null && value !== undefined) {
      eventSheet.getRange(row, col).setValue(value);
    }

  }

}

/*************************************************
 * FIND VENDOR
 *************************************************/
function findVendor(vendorName) {

  const cacheSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.CACHE_SHEET);
  
  if (!cacheSheet) {
    return null;
  }

  const cache = cacheSheet.getDataRange().getValues();

  const search = normalize(vendorName);

  for (let i = 1; i < cache.length; i++) {

    const name = normalize(cache[i][1]);

    if (name === search) {

      return {
        id: cache[i][0],
        email: cache[i][2],
        website: cache[i][3],
        contact: cache[i][4],
        portfolio: cache[i][5]
      };

    }

  }

  return null;

}

/*************************************************
 * FIND VENDOR ROW IN EVENT SHEET
 *************************************************/
function findVendorRow(
  sheet,
  block,
  vendorName
) {

  const cols = getBlockColumns(sheet, block);

  const nameCol = cols.vendor_name;

  const firstVendorRow = block.startRow + 3;
  const lastVendorRow = block.endRow - 1;

  for (let row = firstVendorRow; row <= lastVendorRow; row++) {

    const name = sheet.getRange(row, nameCol).getValue();

    if (normalize(name) === normalize(vendorName)) {
      return row;
    }

  }

  return -1;

}

/*************************************************
 * FIND EMPTY VENDOR ROW
 *************************************************/
function findEmptyVendorRow(sheet, block) {

  const columns = getBlockColumns(sheet, block);

  const nameCol = columns["vendor_name"];

  if (!nameCol) {
    throw new Error("Vendor Name column not found.");
  }

  const firstVendorRow = block.startRow + 3;
  const lastVendorRow = block.endRow - 1;

  for (let row = firstVendorRow; row <= lastVendorRow; row++) {

    const name = sheet.getRange(row, nameCol).getValue();

    if (!name) {
      return row;
    }
  }

  return -1;

}

/*************************************************
 * NORMALIZE
 *
 * Shared string-normalization helper: lowercases,
 * trims, and collapses to snake_case. Used both
 * for vendor-name matching and for turning sheet
 * headers into column keys (see getBlockColumns).
 *************************************************/
function normalize(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[()]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

/*************************************************
 * UPDATE VENDOR
 *
 * Re-syncs an event-sheet vendor row against the
 * Vendor_Cache after the vendor name is entered
 * or changed (called from onEdit and VendorFinder).
 *************************************************/
function updateVendor(sheet, block, row) {

  const cols = getBlockColumns(sheet, block);

  const vendorName = sheet
    .getRange(row, cols["vendor_name"])
    .getValue()
    .toString()
    .trim();

  if (!vendorName) return;

  const vendor = findVendor(vendorName);

  Logger.log(block);
  Logger.log(cols);
  Logger.log(vendorName);
  Logger.log(vendor);

  if (vendor) {

    updateEventVendor(
      sheet,
      block,
      row,
      {
        type: block.type,
        contact: vendor.contact,
        email: vendor.email,
        website: vendor.website
      }
    );

    sheet.getRange(row, cols["vendor_id"]).setValue(vendor.id);
    sheet.getRange(row, cols["vendor_status"]).setValue("Existing");

  } else {

    sheet.getRange(row, cols["vendor_id"]).clearContent();
    sheet.getRange(row, cols["vendor_status"]).setValue("New Vendor");

  }

}
