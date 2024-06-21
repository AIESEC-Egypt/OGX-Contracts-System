function checkApprovals() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "OGX Contracts System"
  );
  var ogx_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APDs");
  var data = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();
  var colIndex = sheet
    .createTextFinder("EP Approved?")
    .matchEntireCell(true)
    .findAll()
    .map((x) => x.getColumn());
  for (let i = 1; i < data.length; i++) {
    sheet
      .getRange(i + 1, colIndex[0] + 3)
      .setValue(
        `=if(AX${i + 1}="","Not Used",IFERROR(IF(MATCH(AZ${
          i + 1
        },APDs!A:A,0),"Yes"),"No"))`
      );
    sheet.getRange(i + 1, colIndex[0] + 4).setValue(`=B${i + 1}`);
    sheet.getRange(i + 1, colIndex[0] + 5).setValue(`=AJ${i + 1}`);
    if (data[i][49] != "") continue;
    console.log(i);

    if (data[i][1] == "OGX Reapproval Agreement") {
      var rowIndex = ogx_sheet
        .createTextFinder(`${data[i][38]}_${data[i][41]}`)
        .matchEntireCell(true)
        .findAll()
        .map((x) => x.getRow());

      if (rowIndex.length > 0) {
        var value = ogx_sheet.getRange(rowIndex, 8).getValue();
        var date = ogx_sheet.getRange(rowIndex, 3).getValue();
        sheet.getRange(i + 1, colIndex).setValue(value);
        sheet.getRange(i + 1, colIndex[0] + 1).setValue(date);
        sheet
          .getRange(i + 1, colIndex[0] + 2)
          .setValue(data[i][38] + "_" + data[i][41]);
      }
    } else {
      var rowIndex = ogx_sheet
        .createTextFinder(`${data[i][14]}_${data[i][24]}`)
        .matchEntireCell(true)
        .findAll()
        .map((x) => x.getRow());
      if (rowIndex.length > 0) {
        var value = ogx_sheet.getRange(rowIndex, 8).getValue();
        var date = ogx_sheet.getRange(rowIndex, 3).getValue();
        sheet.getRange(i + 1, colIndex).setValue(value);
        sheet.getRange(i + 1, colIndex[0] + 1).setValue(date);
        sheet
          .getRange(i + 1, colIndex[0] + 2)
          .setValue(data[i][14] + "_" + data[i][24]);
      }
    }
  }
}

function dataExtraction(query) {
  query = JSON.stringify({ query: query });
  var requestOptions = {
    method: "post",
    payload: query,
    contentType: "application/json",
    headers: {},
  };
  var response = UrlFetchApp.fetch(
    `https://gis-api.aiesec.org/graphql?access_token=${requestOptions["headers"]["access_token"]}`,
    requestOptions
  );
  var recievedDate = JSON.parse(response.getContentText());
  return recievedDate.data.allOpportunityApplication.data;
}

function approvalsUpdating() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APDs");
  var startDate = "01/06/2023";
  var query = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n\t\t\tperson_home_mc:1609\n\t\tdate_approved:{from:\"${startDate}\"}\n\t\t}\n  \n\tpage:1\n    per_page:4000\n\t  \n\t)\n\t{\n  \t\n\t\tdata{\n\t\t\t status\n\n  person{\n\t\t\t\tid\n\t\t\t\tfull_name\n\t\t\t\t\n\t\t\t\temail\n\t\t\t\thome_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n\t\t\t}\n\t\t\topportunity{\n\t\topportunity_duration_type{\n\t\t\t\tduration_type\n\t\t\t}\t\tid\nprogramme {\n\t\t\t\t\tshort_name_display\n\t\t\t\t}\n\t\t\t}\n\t\t\tdate_approved\n\t\t\tslot{\n\t\t\t\tstart_date\n\t\t\t\tend_date\n\t\t\t}\n\t\t\thost_lc{\n\t\tid\n\t\tname\n\t\t\t}\n\t\t\t\n\t\t\t\n\t\t}\n\t\t\n  }\n}`;
  var data = dataExtraction(query);
  var ids = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues();
  ids = ids.flat(1);
  var rows = [];
  for (let i = 0; i < data.length; i++) {
    var searchingID = data[i].person.id + "_" + data[i].opportunity.id;
    var index = ids.indexOf(searchingID);
    if (index == -1) {
      rows.push([
        data[i].person.id + "_" + data[i].opportunity.id,
        data[i].person.full_name,
        data[i].date_approved != null
          ? data[i].date_approved.toString().substring(0, 10)
          : "-",
        data[i].slot.start_date,
        data[i].slot.end_date,
        data[i].person.home_lc.name,
        data[i].opportunity.programme.short_name_display,
        data[i].status,
        data[i].opportunity.opportunity_duration_type.duration_type,
      ]);
    } else {
      var row = [];
      row.push([
        data[i].person.id + "_" + data[i].opportunity.id,
        data[i].person.full_name,
        data[i].date_approved != null
          ? data[i].date_approved.toString().substring(0, 10)
          : "-",
        data[i].slot.start_date,
        data[i].slot.end_date,
        data[i].person.home_lc.name,
        data[i].opportunity.programme.short_name_display,
        data[i].status,
        data[i].opportunity.opportunity_duration_type.duration_type,
      ]);
      sheet.getRange(index + 2, 1, 1, row[0].length).setValues(row);
    }
  }
  if (rows.length > 0)
    sheet
      .getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);
}

function updateBreaks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APDs");
  var query_approval_broken = `query{allOpportunityApplication(\n\t\tfilters:\n\t\t{\n\t\t\tperson_home_mc:1609\n\t\tcreated_at:{from:\"01/06/2023\"}\nstatus:\"approval_broken\"\n\t\t}\n  \n\tpage:1\n    per_page:4000\n\t  \n\t)\n\t{\n  \t\n\t\tdata{\n\t\t\t status\n\n  person{\n\t\t\t\tid\n\t\t\t\tfull_name\n\t\t\t\t\n\t\t\t\temail\n\t\t\t\thome_lc\n        {\n          name\n        }\n        home_mc\n        {\n          name\n        }\n\t\t\t}\n\t\t\topportunity{\n\t\topportunity_duration_type{\n\t\t\t\tduration_type\n\t\t\t}\t\tid\nprogramme {\n\t\t\t\t\tshort_name_display\n\t\t\t\t}\n\t\t\t}\n\t\t\tdate_approved\n\t\t\tslot{\n\t\t\t\tstart_date\n\t\t\t\tend_date\n\t\t\t}\n\t\t\thost_lc{\n\t\tid\n\t\tname\n\t\t\t}\n\t\t\t\n\t\t\t\n\t\t}\n\t\t\n  }\n}`;

  var data = dataExtraction(query_approval_broken);

  for (let i = 0; i < data.length; i++) {
    var searchingID = data[i].person.id + "_" + data[i].opportunity.id;
    var rowIndex = sheet
      .createTextFinder(`${searchingID}`)
      .matchEntireCell(true)
      .findAll()
      .map((x) => x.getRow());
    if (rowIndex.length != 0) {
      sheet.getRange(rowIndex, 8).setValue(data[i].status);
    }
  }
}

function sendNoneEmail() {
  const referenceSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("References");
  var lcs = referenceSheet
    .getRange(2, 1, referenceSheet.getLastRow() - 1, 1)
    .getValues();
  var emails = referenceSheet
    .getRange(
      2,
      1,
      referenceSheet.getLastRow() - 1,
      referenceSheet.getLastColumn()
    )
    .getValues();
  lcs = lcs.flat(1);
  const approvals_sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APDs");
  var approvals_sheet_data = approvals_sheet
    .getRange(
      1,
      1,
      approvals_sheet.getLastRow(),
      approvals_sheet.getLastColumn()
    )
    .getValues();

  for (let i = 0; i < approvals_sheet_data.length; i++) {
    if (
      approvals_sheet_data[i][11] == "" &&
      approvals_sheet_data[i][9] == "None"
    ) {
      var lc_name = approvals_sheet_data[i][5];
      var index = lcs.indexOf(lc_name);
      var ep_id = approvals_sheet_data[i][0].split("_")[0];
      var opp_id = approvals_sheet_data[i][0].split("_")[1];
      var message = `Dear ${emails[index][1]},\nWe hope this email finds you well.\n\nThere's a new approval on EXPA that doesn't have a generated contract.\n\nThe approval details:\nEP ID: ${ep_id}\nOpportunity ID: ${opp_id}\nEP Name: ${approvals_sheet_data[i][1]}\n\nPlease, asap fill the contract system form regarding this EP. As this won't be considered as a real contract so it will lead to under probation & ECB failing submission of this EP.\n\nIf you have any questions don't hesitate to approach us.`;
      if (approvals_sheet_data[i][6] == "GV") {
        var mcvp_resp = mcvp_ogv;
      } else {
        var mcvp_resp = mcvp_ogt;
      }
      MailApp.sendEmail({
        to: `${emails[index][2]}`,
        subject: `OGX Approval Without Generated Contract - ${approvals_sheet_data[i][1]}`,
        body: message,
        cc: `${mcvp_fm},${mcvp_fnl},${mcvp_resp}`,
      });
      approvals_sheet
        .getRange(i + 1, 12)
        .setValue(true)
        .setBackground("green")
        .setFontColor("white");
    }
  }
}

function fillFormulas(sheet) {
  for (let i = 2; i <= sheet.getLastRow(); i++) {
    sheet
      .getRange(i, 10)
      .setFormula(
        `=IFERROR(VLOOKUP(A${i},'OGX Contracts System'!AZ:BB,3,false),"None")`
      );
    sheet
      .getRange(i, 11)
      .setFormula(
        `=IFERROR(VLOOKUP(A${i},'OGX Contracts System'!AZ:BC,4,false),"None")`
      );
  }
}
