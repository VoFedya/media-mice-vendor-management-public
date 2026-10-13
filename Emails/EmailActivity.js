/*************************************************
 * EMAIL ACTIVITY
 *
 * Polls Gmail for vendor replies after an RFQ was
 * sent, and the hourly trigger that drives it.
 *************************************************/

/*************************************************
 * CHECK NEW EMAIL ACTIVITY
 *************************************************/
function checkNewActivity() {

  let checked = 0;
  let updated = 0;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  Logger.log("Running on sheet: " + sheet.getName());

  const blocks = findBlocks(sheet);

  for (const block of blocks) {

    const cols = getBlockColumns(sheet, block);

    const firstVendorRow = block.startRow + 3;
    const lastVendorRow = block.endRow - 1;

    for (let row = firstVendorRow; row <= lastVendorRow; row++) {

      const email =
        sheet.getRange(row, cols.email).getValue();

      const rfqSentNote =
        sheet.getRange(row, cols.email_sent).getNote();
      
      if (!email) continue;
      if (!rfqSentNote) continue;

      let rfqSent;

      try {

        const sentData =
          JSON.parse(rfqSentNote);

        rfqSent =
          new Date(sentData.date);

      } catch (error) {

        rfqSent =
          new Date(rfqSentNote);

      }

      const activityCell =
        sheet.getRange(row, cols.last_activity);

      const lastActivityNote =
        activityCell.getNote();

      let lastActivity = null;

      if (lastActivityNote) {

        const activityData =
          JSON.parse(lastActivityNote);

        lastActivity =
          new Date(activityData.date);

      }

      if (!email) continue;
      if (!rfqSent) continue;

      checked++;

      const threads =
        GmailApp.search(`from:${email}`);

      let latestActivity = null;
      let latestMessage = null;
      let latestThread = null;

      for (const thread of threads) {

        const messages = thread.getMessages();

        for (const message of messages) {

          const date = message.getDate();

          Logger.log({
            messageDate: message.getDate(),
            rfqSent: rfqSent,
            isAfter: message.getDate() > rfqSent,
            from: message.getFrom(),
            subject: message.getSubject()
          });

          // Ignore emails before RFQ
          if (date <= rfqSent) continue;

          // Find newest email after RFQ
          if (!latestActivity || date > latestActivity) {

            latestActivity = date;
            latestMessage = message;
            latestThread = thread;

          }

        }

      }

      if (
        latestActivity &&
        (
          !lastActivity ||
          latestActivity > lastActivity
        )
      ) {

        const subject = latestMessage.getSubject();

        const threadUrl =
          `https://mail.google.com/mail/u/0/#inbox/${latestThread.getId()}`;

        activityCell.setFormula(
          `=HYPERLINK("${threadUrl}","${subject.replace(/"/g, '""')}")`
        );

        activityCell.setNote(
          JSON.stringify({
            date: latestActivity.toISOString(),
            threadId: latestThread.getId()
          })
        );

        activityCell.setBackground("#b7e1cd");

        sheet
          .getRange(row, cols.email_sent)
          .setBackground("#FFE599");

        updated++;

      }

    }

  }

  if (updated === 0) {

    ss.toast(
      `Checked ${checked} vendors.\nNo new activity found.`,
      "Vendor Activity",
      4
    );

  } else {

    ss.toast(
      `Checked ${checked} vendors.\n${updated} new ${updated === 1 ? "activity" : "activities"} found.`,
      "Vendor Activity",
      4
    );

  }

}

/*************************************************
 * HOURLY EMAIL CHECK
 *************************************************/
function hourlyEmailCheck() {

  const hour = new Date().getHours();

  if (hour < 8 || hour >= 18) {
    return;
  }

  checkNewActivity();

}

/*************************************************
 * CREATE HOURLY EMAIL CHECK TRIGGER
 *************************************************/
function createHourlyEmailCheckTrigger() {

  ScriptApp.newTrigger("hourlyEmailCheck")
    .timeBased()
    .everyHours(1)
    .create();

}
