const SHEETID = '1pb9NeXzk4ajlNCmIvkv1LYILgZ5sEK9mxoH8OKMTKZY';
const DOCID = '10uVOJpwinch3B6a6pBG37jCMP7hol4Qsp0X0fjzoemc';
const FOLDERID = '1xiy-doAeleiWJULlv8Wol-4wSiZIC7GR';

function onOpen() {
  SpreadsheetApp.getUi()
.createMenu('Email Out')
.addItem('Send Emails', 'sender')
.addToUi();
}

function sender() {
  const sheet = SpreadsheetApp.openById(SHEETID).getSheetByName('data');
  const temp = DriveApp.getFileById(DOCID);
  const folder = DriveApp.getFolderById(FOLDERID);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  rows.forEach((row, index) => {
      if (row[4] == '') {
      const file = temp.makeCopy(folder);
      const doc = DocumentApp.openById(file.getId());
      const body = doc.getBody();
      data[0].forEach((heading, i) => {
        const header1 = heading.toUpperCase();
        body.replaceText(`{${header1}}`, row[i]);
    })

  doc.setName(row[0] + row[1]);
  const blob = doc.getAs(MimeType.PDF);
  doc.saveAndClose();

  const pdf =  folder.createFile(blob).setName(row[0] + ' ' + row[1] +  '.pdf');
  const email_1 = row[3];
  const email_2 = row[5];
  const subject = row[0] + ' ' + row[1] + ' ' + 'new file created';
  const messageBody = `Hi, ${row[1]} Welcome we've created a PDF for you!`;

  MailApp.sendEmail({
    to: email_1,
    cc: email_2,
    subject: subject,
    htmlBody: messageBody,
    attachments: [blob.getAs(MimeType.PDF)]
  });

  const tempo = sheet.getRange(index + 2, 5, 1, 1);
  tempo.setValue(new Date());
  Logger.log(row);
  file.setTrashed(true);

      }
    })
  }
