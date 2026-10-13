/*************************************************
 * DRIVE MANAGER
 *
 * Drive/Forms side of event creation: the event
 * folder, its copy of the RFQ document templates,
 * and the RFQ intake form.
 *************************************************/

/*************************************************
 * CREATE EVENT FOLDER
 *************************************************/
function createEventFolder(eventName) {

  const root = DriveApp.getFolderById(
    CONFIG.RFQ_ROOT_FOLDER_ID
  );

  const folder = root.createFolder(eventName);

  SpreadsheetApp.getActiveSpreadsheet()
    .getActiveSheet()
    .getRange(EVENT.FOLDER_URL)
    .setFormula(
      `=HYPERLINK("${folder.getUrl()}","📁 Event Folder")`
    );

  return folder;

}

/*************************************************
 * COPY RFQ TEMPLATES
 *************************************************/
function copyRFQTemplates(eventFolder) {

  const source = DriveApp.getFolderById(
    CONFIG.RFQ_TEMPLATE_FOLDER_ID
  );

  const templateFolder =
    eventFolder.createFolder("RFQ Templates");

  const templates = {};

  const files = source.getFiles();

  while (files.hasNext()) {

    const file = files.next();

    const copy = file.makeCopy(
      file.getName(),
      templateFolder
    );

    const vendorType =
      file.getName().replace(/^RFQ - /, "");

    templates[vendorType] = copy.getId();

  }

  return {
    folder: templateFolder,
    templates: templates
  };

}

/*************************************************
 * CREATE RFQ FORM
 *************************************************/
function createRFQForm(eventName, eventFolder) {

  const template = DriveApp.getFileById(
    CONFIG.FORM_TEMPLATE_ID
  );

  const copy = template.makeCopy(
    `${eventName} RFQ`,
    eventFolder
  );

  const form = FormApp.openById(copy.getId());

  form.setDestination(
    FormApp.DestinationType.SPREADSHEET,
    SpreadsheetApp.getActiveSpreadsheet().getId()
  );

  const responseSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheets()
    .find(sheet => sheet.getFormUrl());

  if (responseSheet) {
    responseSheet.hideSheet();
  }

  return {
    id: form.getId(),
    editUrl: form.getEditUrl(),
    url: form.getPublishedUrl(),
  };

}