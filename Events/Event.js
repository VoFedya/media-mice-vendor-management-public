/*************************************************
 * EVENT
 *
 * Reading event metadata and creating new event
 * sheets (spins up the Drive folder, RFQ form,
 * and RFQ template copies via DriveManager).
 *************************************************/

/*************************************************
 * GET EVENT DATA
 *************************************************/
function getEvent(sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()) {

  return {
    name: sheet.getRange(EVENT.NAME).getValue(),
    country: sheet.getRange(EVENT.COUNTRY).getValue(),
    city: sheet.getRange(EVENT.CITY).getValue(),
    date: sheet.getRange(EVENT.DATE).getValue(),
    pm: sheet.getRange(EVENT.PM).getValue(),

    editFormUrl: sheet.getRange(EVENT.EDIT_FORM_URL).getValue(),
    publicFormUrl: sheet.getRange(EVENT.PUBLIC_FORM_URL).getValue(),
    draftEmail: sheet.getRange(EVENT.DRAFT_EMAIL).getValue()
  };

}

/*************************************************
 * CREATE NEW EVENT
 *************************************************/
function createNewEvent() {

  const html = HtmlService
    .createHtmlOutputFromFile("CreateEventDialog")
    .setWidth(420)
    .setHeight(340);

  SpreadsheetApp.getUi().showModalDialog(html, "Create New Event");

}

/*************************************************
 * COPY EVENT SHEET
 *************************************************/
function copyEventSheet(eventName) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const template = ss.getSheetByName(CONFIG.EVENT_TEMPLATE_SHEET);

  const sheet = template.copyTo(ss);

  sheet.setName(eventName);

  const protection = sheet
    .getRange("A:A")
    .protect()
    .setDescription("System-managed Vendor ID");

  protection.setWarningOnly(false);

  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(ss.getNumSheets());

  return sheet;

}

/*************************************************
 * FILL EVENT INFO
 *************************************************/
function fillEventInfo(sheet, data) {

  sheet.getRange(EVENT.NAME).setValue(data.eventName);
  sheet.getRange(EVENT.COUNTRY).setValue(data.country);
  sheet.getRange(EVENT.CITY).setValue(data.city);

  sheet
    .getRange(EVENT.DATE)
    .setValue(formatEventDate(data.startDate, data.endDate));

  sheet.getRange(EVENT.PM).setValue(data.pm);

}

/*************************************************
 * SAVE EVENT RESOURCES
 *************************************************/
function saveEventResources(sheet, rfq, rfqTemplates) {

  sheet
    .getRange(EVENT.EDIT_FORM_URL)
    .setFormula(
      `=HYPERLINK("${rfq.editUrl}","Edit Form")`
    );

  sheet
    .getRange(EVENT.PUBLIC_FORM_URL)
    .setFormula(
      `=HYPERLINK("${rfq.url}","Public Form")`
    );

  sheet
    .getRange(EVENT.DRAFT_EMAIL)
    .setFormula(
      `=HYPERLINK("${rfqTemplates.folder.getUrl()}","RFQ Templates")`
    );

  const configSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Config");

  const sheetId =
    sheet.getSheetId();

  const configData =
    configSheet
      .getRange(2, 1, configSheet.getLastRow() - 1, 1)
      .getValues();

  const rowIndex =
    configData.findIndex(row => row[0] === sheetId);

  if (rowIndex === -1) {
    throw new Error(
      `Event sheet "${eventName}" not found in Config.`
    );
  }

  const configRow = rowIndex + 2;

  configSheet
    .getRange(configRow, 4, 1, 10)
    .setValues([[
      rfqTemplates.templates.Videography || "",
      rfqTemplates.templates.Printing || "",
      rfqTemplates.templates.Livestream || "",
      rfqTemplates.templates.Photography || "",
      rfqTemplates.templates["Call Request"] || "",
      rfqTemplates.templates["Call Follow-up"] || "",
      rfqTemplates.templates.Rejection || "",
      rfqTemplates.templates.Selection || "",
      rfqTemplates.templates.Other || "",
      rfqTemplates.templates.AudioVisual || ""
    ]]);

}

/*************************************************
 * CREATE NEW EVENT FROM DIALOG
 *************************************************/
function createNewEventFromDialog(data) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const eventName = data.eventName.trim();

  if (!eventName) {
    ui.alert("Event name cannot be empty.");
    return;
  }

  if (ss.getSheetByName(eventName)) {
    ui.alert("A sheet with this name already exists.");
    return;
  }

  const sheet = copyEventSheet(eventName);

  fillEventInfo(sheet, data);

  const folder = createEventFolder(eventName);
  const rfqTemplates = copyRFQTemplates(folder);
  const rfq = createRFQForm(eventName, folder);

  saveFormConfig(sheet.getSheetId(), rfq.id);

  saveEventResources(sheet, rfq, rfqTemplates);

  ui.alert(`"${eventName}" created successfully.`);

}


/*************************************************
 * FORMAT EVENT DATE
 *************************************************/
function formatEventDate(start, end) {

  const tz = Session.getScriptTimeZone();

  const startDate = new Date(start);

  if (!end || start === end) {
    return Utilities.formatDate(startDate, tz, "d MMM yyyy");
  }

  const endDate = new Date(end);

  const sameYear =
    startDate.getFullYear() === endDate.getFullYear();

  const sameMonth =
    sameYear &&
    startDate.getMonth() === endDate.getMonth();

  if (sameMonth) {
    return `${Utilities.formatDate(startDate, tz, "d")}–${Utilities.formatDate(endDate, tz, "d MMM yyyy")}`;
  }

  if (sameYear) {
    return `${Utilities.formatDate(startDate, tz, "d MMM")} – ${Utilities.formatDate(endDate, tz, "d MMM yyyy")}`;
  }

  return `${Utilities.formatDate(startDate, tz, "d MMM yyyy")} – ${Utilities.formatDate(endDate, tz, "d MMM yyyy")}`;

}
