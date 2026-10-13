/*************************************************
 * ON OPEN
 *************************************************/
function onOpen() {

  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Event Tools")
    .addItem("➕ Create Event", "createNewEvent")
    .addSeparator()
    .addItem("🗂 Organize Tabs", "organizeTabs")
    .addSeparator()
    .addItem("🔎 Find Vendors", "openVendorFinder")
    .addSeparator()
    .addItem("📧 Send Emails", "sendRFQs")
    .addItem("📬 Check Replies", "checkNewActivity")
    .addSeparator()
    .addItem("📊 Push to Evaluation", "pushToEvaluation")
    .addToUi();

  ui.createMenu("Block Manager")
    .addItem("🧱 Add Vendor Block", "showAddVendorCategoryDialog")
    .addItem("🗑️ Delete Current Block", "deleteCurrentBlock")
    .addSeparator()
    .addItem("⬇️ Add Vendor Rows", "addVendorRows")
    .addItem("➡️ Add Price Column", "addPriceColumn")
    .addToUi();

  ui.createMenu("MDM Tools")
    .addItem("📥 Pull Evaluation Scores", "pullEvaluationScores")
    .addItem("📤 Push Event to MDM", "pushEventToMDM")
    .addItem("📤 Push Vendors to MDM Queue", "submitToMDM")
    .addItem("📤 Push Pricing to MDM", "pushPricingToMDM")
    .addToUi();

}

function onEdit(e) {

  const sheet = e.range.getSheet();

  if (CONFIG.SYSTEM_SHEETS.includes(sheet.getName())) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  const block = findBlocks(sheet).find(block =>
    row >= block.startRow + 3 &&
    row <= block.endRow - 1
  );

  if (!block) return;

  const cols = getBlockColumns(sheet, block);

  if (col !== cols["vendor_name"]) return;

  updateVendor(sheet, block, row);

}

/*************************************************
 * ADD VENDOR CATEGORY DIALOG
 *************************************************/
function showAddVendorCategoryDialog() {
  const html = HtmlService
    .createHtmlOutputFromFile("AddVendorCategory")
    .setWidth(320)
    .setHeight(180);

  SpreadsheetApp
    .getUi()
    .showModalDialog(html, "Add Vendor Category");
}
