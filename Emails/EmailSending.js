/*************************************************
 * EMAIL SENDING
 *
 * Sends RFQ/communication emails to vendors marked
 * "send" on the event sheet, either as a reply on
 * an existing Gmail thread or as a new message.
 *************************************************/

/*************************************************
 * SEND EMAILS
 *************************************************/
function sendRFQs() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  let sent = 0;

  const blocks = findBlocks(sheet);

  for (const block of blocks) {

    const cols = getBlockColumns(sheet, block);

    const firstVendorRow = block.startRow + 3;
    const lastVendorRow = block.endRow - 1;

    for (let row = firstVendorRow; row <= lastVendorRow; row++) {

      const send =
        sheet.getRange(row, cols.send_email).getValue();

      const emailType =
        sheet.getRange(row, cols.email_type).getValue();

      const vendor =
        sheet.getRange(row, cols.vendor_name).getValue();

      const email =
        sheet.getRange(row, cols.email).getValue();

      if (!send || !email) continue;

      if (!emailType) {
        SpreadsheetApp.getUi().alert(
          `No Email Type selected for ${vendor}.`
        );
        return;
      }

      const activityNote =
        sheet.getRange(row, cols.last_activity).getNote();

      let threadId = null;

      if (activityNote) {

        try {

          const activityData =
            JSON.parse(activityNote);

          threadId =
            activityData.threadId;

        } catch (error) {

          threadId = null;

        }

      }

      try {

        const emailData =
          buildEmail(sheet, block, row, vendor, emailType);

        if (threadId) {

          const thread =
            GmailApp.getThreadById(threadId);

          thread.replyAll("", {
            htmlBody: emailData.html,
            from: EMAIL.VENDOR
          });

          const threadUrl =
            `https://mail.google.com/mail/u/0/#inbox/${threadId}`;

          sheet
            .getRange(row, cols.email_sent)
            .setFormula(
              `=HYPERLINK("${threadUrl}","Sent")`
            );

          sheet
            .getRange(row, cols.email_sent)
            .setNote(
              JSON.stringify({
                date: new Date().toISOString(),
                threadId: threadId
              })
            );

        } else {

          GmailApp.sendEmail(
            email,
            emailData.subject,
            "",
            {
              htmlBody: emailData.html,
              from: EMAIL.VENDOR
            }
          );

          saveSentEmailLink(
            sheet,
            cols,
            row,
            email,
            emailData.subject
          );

        }

        sent++;

        sheet
          .getRange(row, cols.email_sent)
          .setBackground("#b7e1cd");

        sheet
          .getRange(row, cols.last_activity)
          .setBackground("#FFE599");

        sheet
          .getRange(row, cols.send_email)
          .setValue(false);

      } catch (error) {

        Logger.log(
          `Failed to send to ${vendor} (${email}): ${error}`
        );

        sheet
          .getRange(row, cols.email_sent)
          .setValue("FAILED");

        SpreadsheetApp
          .getUi()
          .alert(error.message);

        return;

      }

    }

  }

  SpreadsheetApp.getUi().alert(
    `${sent} email${sent === 1 ? "" : "s"} sent successfully.`
  );

}

/*************************************************
 * BUILD EMAIL
 *
 * Loads the right template (RFQ or communication)
 * and fills in the {{PLACEHOLDER}} fields for one
 * vendor/row.
 *************************************************/
function buildEmail(sheet, block, row, vendor, emailType) {

  let template;

  if (emailType === "RFQ") {

    template =
      getRFQTemplate(block.type);

    validateRFQTemplate(template);

  } else {

    template =
      getCommunicationTemplate(emailType);

    validateCommunicationTemplate(template);

  }

  const eventName =
    sheet.getRange(EVENT.NAME).getValue();

  const country =
    sheet.getRange(EVENT.COUNTRY).getValue();

  const city =
    sheet.getRange(EVENT.CITY).getValue();

  const eventDate =
    sheet.getRange(EVENT.DATE).getValue();

  const pm =
    sheet.getRange(EVENT.PM).getValue();

  const formLink =
    sheet
      .getRange(EVENT.PUBLIC_FORM_URL)
      .getFormula()
      .match(/HYPERLINK\("([^"]+)"/i)?.[1] || "";

  let subject = template.subject;
  let html = template.html;

  subject = subject
    .replace(/{{EVENT}}/g, eventName)
    .replace(/{{CITY}}/g, city)
    .replace(/{{VENDOR}}/g, vendor)
    .replace(/{{EVENT_DATE}}/g, eventDate);

  html = html
    //.replace(/{{CONTACT}}/g, contact)
    .replace(/{{VENDOR}}/g, vendor)
    .replace(/{{EVENT}}/g, eventName)
    .replace(/{{CITY}}/g, city)
    .replace(/{{COUNTRY}}/g, country)
    .replace(/{{EVENT_DATE}}/g, eventDate)
    .replace(/{{PM}}/g, pm)
    .replace(/{{FORM_LINK}}/g, formLink);

  return {
    subject,
    html
  };

}

/*************************************************
 * SAVE SENT EMAIL LINK
 *************************************************/
function saveSentEmailLink(
  sheet,
  cols,
  row,
  email,
  subject
) {

  // Give Gmail a moment to create the thread
  Utilities.sleep(1000);

  const threads = GmailApp.search(
    `to:${email} subject:"${subject}" newer_than:1d`,
    0,
    1
  );

  if (threads.length === 0) {

    sheet
      .getRange(row, cols.email_sent)
      .setValue("Sent");

    return;

  }

  const thread = threads[0];

  const url =
    `https://mail.google.com/mail/u/0/#inbox/${thread.getId()}`;

  const cell =
    sheet.getRange(row, cols.email_sent);

  cell.setFormula(
    `=HYPERLINK("${url}","Sent")`
  );

  cell.setNote(
    JSON.stringify({
      date: new Date().toISOString(),
      threadId: thread.getId()
    })
  );

}
