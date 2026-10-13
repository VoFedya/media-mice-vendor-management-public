/*************************************************
 * ORGANIZE TABS
 *************************************************/
function organizeTabs() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Visible sheets in desired order
  const order = [
    "User_Guide",
    "README",
    "Config",
    "Event_Template",
    "Printing_Block",
    "Videography_Block",
    "Livestream_Block",
    "Photography_Block",
    "AV_Block",
    "Other_Block",
    "Vendor_Search",
    "Vendor_Cache"
  ];

  let position = 1;

  // Move fixed sheets first
  order.forEach(name => {

    const sheet = ss.getSheetByName(name);

    if (!sheet) return;

    sheet.showSheet();

    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(position++);

  });

  // Hide system sheets
  const hiddenSheets = [
    "README",
    "Config",
    "Event_Template",
    "Printing_Block",
    "Videography_Block",
    "Livestream_Block",
    "AudioVisual_Block",
    "Photography_Block",
    "Other_Block",
    "Vendor_Cache",
    "AV_Block",
  ];

  hiddenSheets.forEach(name => {

    const sheet = ss.getSheetByName(name);

    if (sheet) {
      sheet.hideSheet();
    }

  });

  // Hide Form Responses sheets
  ss.getSheets()
    .filter(sheet => sheet.getName().startsWith("Form Responses"))
    .forEach(sheet => sheet.hideSheet());

}