/*************************************************
 * OPEN VENDOR FINDER
 *************************************************/
function openVendorFinder() {

  const html = HtmlService
    .createHtmlOutputFromFile("Vendor_Finder")
    .setTitle("Vendor Finder");

  SpreadsheetApp.getUi().showSidebar(html);

}

/*************************************************
 * GET VENDOR SEARCH DATA
 *************************************************/
function getVendorSearch() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.VENDOR_SEARCH);

  const data = sheet
    .getDataRange()
    .getValues();

  const vendors = new Map();

  for (
    let i = CONFIG.ROW.VENDOR_SEARCH.FIRST_DATA - 1;
    i < data.length;
    i++
  ) {

    const id = data[i][0];

    if (!id) continue;

    if (!vendors.has(id)) {

      vendors.set(id, {

        id: id,
        name: data[i][1],
        types: new Set(),

        evaluated: data[i][5],
        selected: data[i][6],

        email: data[i][8],
        website: data[i][9],
        contact: data[i][10],
        portfolio: data[i][11],

        coverage: [],
        eventBlocks: []

      });

    }

    vendors.get(id).types.add(data[i][2]);

    vendors.get(id).coverage.push({

      country: data[i][3],
      city: data[i][4],
      evaluated: data[i][5],
      selected: data[i][6]

    });

  }

  const blocks = findBlocks(
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  );

  blocks.forEach(block => {

    const cols = getBlockColumns(
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
      block
    );

    const firstVendorRow = block.startRow + 3;
    const lastVendorRow = block.endRow - 1;

    for (
      let row = firstVendorRow;
      row <= lastVendorRow;
      row++
    ) {

      const vendorId =
        String(
          SpreadsheetApp
            .getActiveSpreadsheet()
            .getActiveSheet()
            .getRange(row, cols.vendor_id)
            .getValue()
        ).trim();

      if (!vendorId) continue;

      if (vendors.has(vendorId)) {

        vendors
          .get(vendorId)
          .eventBlocks
          .push(block.type);

      }

    }

  });

  const result = [...vendors.values()];

  result.forEach(v => {
    v.type = [...v.types].join("; ");
    delete v.types;
  });

  return result;

}

/*************************************************
 * GET CURRENT CONTEXT
 *************************************************/
function getCurrentContext() {

  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveCell().getRow();
  
  Logger.log(sheet.getName());
  Logger.log(row);
  Logger.log(findBlocks(sheet));

  const block = findBlocks(sheet).find(
    block => row >= block.startRow && row <= block.endRow
  );

  const firstVendorRow = block ? block.startRow + 3 : 0;
  const lastVendorRow = block ? block.endRow - 1 : 0;

  const valid =
    !CONFIG.SYSTEM_SHEETS.includes(sheet.getName()) &&
    !!block &&
    row >= firstVendorRow &&
    row <= lastVendorRow;

  let canAdd = false;

  if (valid) {

    const columns = getBlockColumns(sheet, block);

    const vendorName =
      sheet.getRange(row, columns.vendor_name).getValue();

    canAdd = vendorName === "";

  }
  Logger.log({
    sheet: sheet.getName(),
    row,
    block,
    valid,
    canAdd
  });

  return {
    sheet: sheet.getName(),
    row: row,
    valid: valid,
    canAdd: canAdd,
    blockType: block?.type
  };

}


/*************************************************
 * ADD VENDOR TO CURRENT ROW
 *************************************************/
function addVendorToCurrentRow(vendorName) {

  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveCell().getRow();

  if (CONFIG.SYSTEM_SHEETS.includes(sheet.getName())) {
    throw new Error("Select an event sheet.");
  }

  const block = findBlocks(sheet).find(
    block => row >= block.startRow && row <= block.endRow
  );

  if (!block) {
    throw new Error("Select a vendor row first.");
  }

  const columns = getBlockColumns(sheet, block);

  // Don't overwrite existing vendors
  if (sheet.getRange(row, columns.vendor_name).getValue() !== "") {
    throw new Error("This row already contains a vendor.");
  }

  // Write vendor name
  sheet
    .getRange(row, columns.vendor_name)
    .setValue(vendorName);

  // Autofill vendor information
  updateVendor(sheet, block, row);;

}


/*************************************************
 * GET VENDOR FILTERS
 *************************************************/
function getVendorFilters() {

  const vendors = getVendorSearch();

  const types = new Set();
  const countries = new Set();
  const cities = new Set();

  const countryCities = {};

  vendors.forEach(v => {

    getVendorTypes(v.type).forEach(type => {

      types.add(type);

    });

    v.coverage.forEach(c => {

      countries.add(c.country);
      cities.add(c.city);

      if (!countryCities[c.country]) {
        countryCities[c.country] = new Set();
      }

      countryCities[c.country].add(c.city);

    });

  });

  const result = {

    types: [...types].sort(),
    countries: [...countries].sort(),
    cities: [...cities].sort(),
    countryCities: {}

  };

  for (const country in countryCities) {

    result.countryCities[country] =
      [...countryCities[country]].sort();

  }

  return result;

}

/*************************************************
 * GET VENDOR TYPES
 *************************************************/
function getVendorTypes(typeString) {

  return String(typeString || "")
    .split(";")
    .map(type => type.trim())
    .filter(type => type !== "");

}