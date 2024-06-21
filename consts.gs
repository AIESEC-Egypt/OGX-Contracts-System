const referenceSheet =
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reference");
const referenceSheetData = referenceSheet
  .getRange(1, 1, referenceSheet.getLastRow(), referenceSheet.getLastColumn())
  .getValues();

const mcvp_ogv = "omar.otify@aiesec.net";
const mcvp_ogt = "m.abualewa@aiesec.org.eg";
const mcvp_fnl = "o.saifelnasr@aiesec.org.eg";
const mcvp_fm = "y.khaled@aiesec.org.eg";
const mcvp_igv = "m.alswaf@aiesec.org.eg";
const mcvp_igt = "t.maria@aiesec.org.eg";

// LC Codes from Expa padded with 0 on the left
const lcMap = {};

const ecbSheetsMap = {};

const dateFormat = "yyyyddMM";
