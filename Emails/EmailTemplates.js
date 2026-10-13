/*************************************************
 * EMAIL TEMPLATES
 *
 * Loads RFQ/communication template Docs (per
 * event, per vendor type) from Config, parses out
 * their SUBJECT/BODY sections, and validates that
 * required {{PLACEHOLDER}} fields are present.
 *************************************************/

/*************************************************
 * GET RFQ TEMPLATE
 *************************************************/
function getRFQTemplate(vendorType) {

  const configSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Config");

  const sheetId =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getActiveSheet()
      .getSheetId();

  const events =
    configSheet
      .getRange(2, 1, configSheet.getLastRow() - 1, 13)
      .getValues();

  const eventRow =
    events.find(row => row[0] === sheetId);

  if (!eventRow) {
    throw new Error(
      `Event sheet "${sheetId}" not found in Config.`
    );
  }

  const templateColumns = {
    "videography": 3,
    "printing": 4,
    "livestream": 5,
    "photography": 6,
    "other": 11,
    "audiovisual": 12
  };

  const columnIndex = templateColumns[vendorType.toLowerCase()];

  if (columnIndex === undefined) {
    throw new Error(
      `No RFQ template column configured for vendor type "${vendorType}".`
    );
  }

  const templateId = eventRow[columnIndex];

  if (!templateId) {
    throw new Error(
      `No RFQ template copy found for "${vendorType}" for event sheet "${sheetId}".`
    );
  }

  const url =
    `https://docs.google.com/document/d/${templateId}/export?format=html`;

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    }
  });

  return parseRFQTemplate(
    response.getContentText()
  );

}

/*************************************************
 * GET COMMUNICATION TEMPLATE
 *************************************************/
function getCommunicationTemplate(emailType) {

  const configSheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Config");

  const sheetId =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getActiveSheet()
      .getSheetId();

  const data =
    configSheet
      .getRange(
        1,
        1,
        configSheet.getLastRow(),
        configSheet.getLastColumn()
      )
      .getValues();

  const headers = data[0];

  const eventRow =
    data
      .slice(1)
      .find(row => row[0] === sheetId);

  if (!eventRow) {
    throw new Error(
      `Event "${eventName}" not found in Config.`
    );
  }

  const columnIndex =
    headers.indexOf(emailType);

  if (columnIndex === -1) {
    throw new Error(
      `No template column configured for "${emailType}".`
    );
  }

  const templateId =
    eventRow[columnIndex];

  if (!templateId) {
    throw new Error(
      `No template copy found for "${emailType}" for event "${eventName}".`
    );
  }

  const url =
    `https://docs.google.com/document/d/${templateId}/export?format=html`;

  const response =
    UrlFetchApp.fetch(url, {
      headers: {
        Authorization:
          "Bearer " + ScriptApp.getOAuthToken()
      }
    });

  return parseRFQTemplate(
    response.getContentText()
  );

}

/*************************************************
 * STRIP HTML
 *************************************************/
function stripHtml(html) {

  return html
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/\s+/g, " ")
    .trim();

}

/*************************************************
 * PARSE RFQ TEMPLATE
 *************************************************/
function parseRFQTemplate(html) {

  const subjectHeading =
    html.search(/<h1[^>]*>\s*<span[^>]*>\s*SUBJECT\s*<\/span>\s*<\/h1>/i);

  const bodyHeading =
    html.search(/<h1[^>]*>\s*<span[^>]*>\s*BODY\s*<\/span>\s*<\/h1>/i);

  if (subjectHeading === -1) {
    throw new Error('Template is missing the "SUBJECT" heading.');
  }

  if (bodyHeading === -1) {
    throw new Error('Template is missing the "BODY" heading.');
  }

  const subjectSection =
    html.substring(subjectHeading, bodyHeading);

  const paragraphs =
    [...subjectSection.matchAll(/<p[^>]*>([\s\S]*?)<\/p>/g)];

  const subjectMatch =
    paragraphs.find(p => stripHtml(p[1]) !== "");

  if (!subjectMatch) {
    throw new Error("Subject line not found.");
  }

  let body =
    html.substring(bodyHeading)
        .replace(/^.*?<\/h1>/s, "")
        .replace(/<\/body>[\s\S]*$/i, "");

  // Convert Google Docs indented lists into nested lists
  try {

    body = body.replace(
      /(<ul[^>]*lst-kix_[^"]*-0[^>]*>)([\s\S]*?)(<\/ul>)\s*(<ul[^>]*lst-kix_[^"]*-3[^>]*>[\s\S]*?<\/ul>)/gi,

      function(match, open, content, close, nested) {

        const lastLi =
          content.match(/([\s\S]*)(<li[^>]*>[\s\S]*?<\/li>)\s*$/i);

        if (!lastLi) return match;

        return (
          open +
          lastLi[1] +
          lastLi[2].replace(
            /<\/li>\s*$/i,
            nested + "</li>"
          ) +
          close
        );

      }
    );

  } catch (error) {

    Logger.log(
      "Nested list formatting skipped: " + error
    );

  }

  body = cleanGoogleLinks(body);

  Logger.log(body);

  return {

    subject: stripHtml(subjectMatch[1]),

    html: body

  };

}
/*************************************************
 * CLEAN GOOGLE LINKS
 *************************************************/
function cleanGoogleLinks(html) {

  return html.replace(
    /https:\/\/www\.google\.com\/url\?q=([^&"]+)[^"]*/g,
    (_, url) => decodeURIComponent(url)
  );

}

/*************************************************
 * VALIDATE RFQ TEMPLATE
 *************************************************/
function validateRFQTemplate(template) {

  const subjectFields = [
    //"{{EVENT}}",
    //"{{CITY}}",
    "{{VENDOR}}"
  ];

  const bodyFields = [
    "{{VENDOR}}",
    //"{{EVENT}}",
    //"{{EVENT_DATE}}"
  ];

  for (const field of subjectFields) {

    if (!template.subject.includes(field)) {

      throw new Error(
        `RFQ Template Error\n\nThe SUBJECT must contain "${field}".`
      );

    }

  }

  for (const field of bodyFields) {

    if (!template.html.includes(field)) {

      throw new Error(
        `RFQ Template Error\n\nThe BODY must contain "${field}".`
      );

    }

  }

}

/*************************************************
 * VALIDATE COMMUNICATION TEMPLATE
 *************************************************/
function validateCommunicationTemplate(template) {

  const fields = [
    "{{VENDOR}}",
    "{{EVENT}}"
  ];

  for (const field of fields) {

    if (
      !template.subject.includes(field) &&
      !template.html.includes(field)
    ) {

      throw new Error(
        `Communication Template Error\n\nThe template must contain "${field}".`
      );

    }

  }

}
