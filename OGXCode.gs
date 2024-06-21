function createContractOGX() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "OGX Contracts System"
  );
  const sheetData = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();
  const rowIndex = sheet.getLastRow();
  var sendCol = sheet
    .createTextFinder("Email Sent?")
    .matchEntireCell(true)
    .findAll()
    .map((x) => x.getColumn());
  if (sheet.getRange(rowIndex, sendCol).getValue() == true) return;
  if (sheetData[rowIndex - 1][1] == "OGX Reapproval Agreement") {
    var pairs = referenceSheet
      .createTextFinder(`Reapproval`)
      .matchEntireCell(true)
      .findAll()
      .map((x) => [x.getRow(), x.getColumn()]);
    var epName = sheet.getRange(rowIndex, 26).getValue();
  } else {
    var pairs = referenceSheet
      .createTextFinder(`OGX`)
      .matchEntireCell(true)
      .findAll()
      .map((x) => [x.getRow(), x.getColumn()]);
    var epName = sheet.getRange(rowIndex, 3).getValue();
  }
  const contractType = sheet.getRange(rowIndex, 2).getValue();

  const contractIDIndex = referenceSheet
    .createTextFinder(contractType)
    .matchEntireCell(true)
    .findAll()
    .map((x) => x.getRow());
  const contractID = referenceSheet.getRange(contractIDIndex, 2).getValue();

  const template = DriveApp.getFileById(contractID);
  const name = `AIESEC in Egypt - ${contractType} - ${epName}`;
  const newFile = template.makeCopy(name, folder);
  console.log(newFile.getUrl());
  const doc = DocumentApp.openById(newFile.getId());
  const docBody = doc.getBody();

  var emails = [];
  for (let i = 0; i < pairs.length; i++) {
    if (
      !referenceSheetData[pairs[i][0] - 1][pairs[i][1]].includes(
        "AIESEC Representative Email"
      )
    ) {
      var colIndex = sheet
        .createTextFinder(referenceSheetData[pairs[i][0] - 1][pairs[i][1]])
        .matchEntireCell(true)
        .findAll()
        .map((x) => x.getColumn());
      var replaced = referenceSheetData[pairs[i][0] - 1][pairs[i][1] + 1];
      var value = sheetData[rowIndex - 1][colIndex - 1];
      if (replaced.toString().includes("date")) {
        var value = Utilities.formatDate(
          new Date(value),
          "GMT+3",
          "dd/MM/yyyy"
        );
      }
      if (
        referenceSheetData[pairs[i][0] - 1][pairs[i][1]] == "Reference Code"
      ) {
        var indices = referenceSheet
          .createTextFinder(`${lc}`)
          .matchEntireCell(true)
          .findAll()
          .map((x) => [x.getRow(), x.getColumn() + 1]);
        var lcCode = lcMap[`${lc}`];
        var date = Utilities.formatDate(new Date(), "GMT+3", dateFormat);
        var value =
          lcCode + epId + date + Math.floor(Math.random() * 100000 + 1);
        sheet.getRange(rowIndex, colIndex).setValue(value);
      }
      if (
        referenceSheetData[pairs[i][0] - 1][pairs[i][1]].includes(
          "AIESEC in Egypt – Local Committee Branch"
        )
      ) {
        var lc = sheetData[rowIndex - 1][colIndex - 1];
      }
      if (referenceSheetData[pairs[i][0] - 1][pairs[i][1]].includes("EP ID"))
        var epId = sheetData[rowIndex - 1][colIndex - 1].toString().slice(-3);
      docBody.replaceText(replaced, value);
      if (
        referenceSheetData[pairs[i][0] - 1][pairs[i][1]].includes(
          "Exchange Participant Email"
        )
      ) {
        emails.push(sheetData[rowIndex - 1][colIndex - 1]);
      }
    } else {
      emails.push(
        sheetData[rowIndex - 1][
          parseInt(referenceSheetData[pairs[i][0] - 1][pairs[i][1] + 1]) - 1
        ]
      );
    }
  }
  doc.saveAndClose();

  MailApp.sendEmail({
    to: `${emails.join(",")}`,
    subject: `${name}`,
    body: `Hello ${epName},\nGreeting from AIESEC in Egypt.\n\nYou can find your copy of the contract that should be signed in the next few days with your manager.\n\nThank you.\n\nPlease, dont' reply. This is an automated mail.`,
    attachments: [newFile.getAs(MimeType.PDF)],
  });
  sheet
    .getRange(rowIndex, sendCol)
    .setValue(true)
    .setBackground("green")
    .setFontColor("white");
}
