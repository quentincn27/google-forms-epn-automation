function printAndDeleteCells() {

  const spreadsheetId = 'WRITE THE SHEETS URL';
  const statsSheetName = 'Statistiques';
  const timezone = 'Europe/Paris';

  const today = new Date();
  const day = today.getDay(); // 0 = Sunday, 6 = Saturday

  // Cancel execution on weekends
  if (day === 0 || day === 6) {
    Logger.log('Weekend: no email sent');
    return;
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  // Cancel execution on public holidays
  if (isPublicHoliday(today, spreadsheet)) {
    Logger.log('Public holiday: no email sent');
    return;
  }

  const sheet = spreadsheet.getSheetByName(statsSheetName);
  if (!sheet) {
    Logger.log('Statistics sheet not found');
    return;
  }

  // Format current date
  const formattedDate = Utilities.formatDate(today, timezone, 'dd/MM/yyyy');

  // PDF generation
  const url =
    `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
    `format=pdf&` +
    `size=A4&` +
    `portrait=false&` +
    `fitw=true&` +
    `gridlines=false&` +
    `sheetnames=false&` +
    `gid=${sheet.getSheetId()}&` +
    `range=A1:F25`;

  try {
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + token
      }
    });

    const blob = response.getBlob()
      .setName(`EPN statistics - ${formattedDate.replace(/\//g, '-')}.pdf`);

    // Email sent to supervisor
    GmailApp.sendEmail(
      'WRITE THE MAIL',
      `EPN statistics - ${formattedDate}`,
      `Hello WRITE THE NAME,

Please find attached the EPN attendance statistics for today.

Best regards,
WRITE THE NAME`,
      {
        attachments: [blob],
        replyTo: 'WRITE THE MAIL'
      }
    );

    // Confirmation email to myself
    GmailApp.sendEmail(
      'WRITE THE MAIL',
      `Confirmation – EPN statistics ${formattedDate}`,
      `The EPN statistics for ${formattedDate} have been successfully sent.`,
      {
        attachments: [blob]
      }
    );

    // Cleanup AFTER sending
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }

    SpreadsheetApp.flush();

  } catch (error) {
    Logger.log('Error: ' + error.toString());
  }
}

function isPublicHoliday(date, spreadsheet) {

  const sheet = spreadsheet.getSheetByName('JoursFeries');
  if (!sheet) return false;

  const dates = sheet.getRange('A:A').getValues().flat();

  return dates.some(d =>
    d instanceof Date &&
    d.toDateString() === date.toDateString()
  );
}
