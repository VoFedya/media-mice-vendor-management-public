/*************************************************
 * DEBUG: EMAIL TESTS
 *
 * Manual, ad-hoc helpers for testing Gmail
 * sending/aliases from the Apps Script editor.
 * Not wired into any menu or trigger.
 * NOTE: contains a hardcoded personal test address
 * — update or remove before relying on this again.
 *************************************************/

function testGroupEmail() {

  const groupEmail = "test@example.com";
  const recipient = "test@example.com";

  GmailApp.sendEmail(
    recipient,
    "Test from Vendor Group",
    "This is a test email.",
    {
      htmlBody: "<p>This is a test email.</p>",
      from: groupEmail
    }
  );

}

function testAliases() {
  Logger.log(GmailApp.getAliases());
}
