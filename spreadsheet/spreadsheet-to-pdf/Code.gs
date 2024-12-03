async function discoverFields(config) {
  const fileId = config.pdfTemplate.id; // Please set the file ID of the PDF file on Google Drive.
  const blob = DriveApp.getFileById(fileId).getBlob();
  const PF = new PdfForm();
  try {
    const values = await PF.getValues(blob);

    return values.map((field) => ({
      ref: field.ref,
      name: field.name,
      type: field.type,
      defaultValue:
        typeof field.value === "string" ? field.value.trim() : field.value,
    }));
  } catch (error) {
    console.error(error);
    throw error;
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu("Pdf Generator")
    .addItem("Configure", "showConfigModal")
    .addItem("Synchroniser", "synchronize")
    .addItem("Générer les lignes selectionnées", "generateSelectedLines")
    .addToUi();
}

function getIndexesOfSelectedRows(rangeList, bounds) {
  let selectedRowsIndexes = [];
  const ranges = rangeList.getRanges();
  loopRange: for (let i in ranges) {
    let a1Notation = ranges[i].getA1Notation();
    if (!a1Notation.includes(":")) {
      const row = parseInt(a1Notation.replace(/[A-Z]+/, ""));
      a1Notation = `A${row}:${row}`;
    }
    const tmp = a1Notation
      .split(":")
      .map((val) => parseInt(val.replace(/[A-Z]+/, "")));
    const values = SpreadsheetApp.getActiveSheet()
      .getRange(a1Notation)
      .getValues();
    if (tmp[0] > bounds.row) break loopRange;
    for (let j = tmp[0]; j < tmp[1] + 1; j++) {
      if (j > bounds.row) break loopRange;
      if (values[j - tmp[0]][0] === "OK") continue;
      selectedRowsIndexes.push(j);
    }
  }
  return selectedRowsIndexes;
}

function getDataBounds() {
  const dataRange = SpreadsheetApp.getActiveSheet().getDataRange();
  return { col: dataRange.getLastColumn(), row: dataRange.getLastRow() };
}

function indexToColumn(index) {
  // Validate index size
  const maxIndex = 18278;
  if (index > maxIndex) {
    throw new Error(`index cannot be greater than ${maxIndex} (column ZZZ)`);
  }

  // Get column from index
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  if (index > 26) {
    const letterA = indexToColumn(Math.floor((index - 1) / 26));
    const letterB = indexToColumn(index % 26);
    return letterA + letterB;
  } else {
    if (index == 0) {
      index = 26;
    }
    return alphabet[index - 1];
  }
}

function generateSelectedLines() {
  const activeRangeList = SpreadsheetApp.getActiveRangeList();
  const bounds = getDataBounds();
  const rowIndexes = getIndexesOfSelectedRows(activeRangeList, bounds);
  convertRows(rowIndexes, bounds);
}

function synchronize() {
  const bounds = getDataBounds();
  const rangeToSync = `A4:A${bounds.row}`;
  const states = SpreadsheetApp.getActiveSheet()
    .getRange(rangeToSync)
    .getValues();
  const rowIndexes = [];
  for (let i = 0; i < states.length; i++) {
    if (states[i][0] !== "OK") {
      rowIndexes.push(4 + i);
    }
  }
  convertRows(rowIndexes, bounds);
}

async function convertRows(rowIndexes, bounds) {
  const ui = SpreadsheetApp.getUi();
  if (!rowIndexes.length) return ui.alert("Aucun document à générer");

  const confirm = ui.alert(
    `Êtes-vous sûr de vouloir générer ${
      rowIndexes.length
    } documents? Lignes ${rowIndexes.join(", ")}`,
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm == ui.Button.OK) {
    const results = await processRows(rowIndexes, bounds);
    ui.alert(
      `${results.succeed.length} documents générés.\n${
        results.failed.length
      } documents en erreur (Lignes ${results.failed.join(", ")}).`
    );
  }
}

function getColumnsMapping(bounds) {
  const maxCol = indexToColumn(bounds.col);
  const mappingsValues = SpreadsheetApp.getActiveSheet()
    .getRange(`B2:${maxCol}2`)
    .getValues()[0];
  return mappingsValues.reduce((acc, ref, index) => {
    if (ref) {
      acc[index + 2] = ref;
    }
    return acc;
  }, {});
}

async function processRows(rowIndexes, bounds) {
  const columns = getColumnsMapping(bounds);
  return Promise.allSettled(
    rowIndexes.map((rowIndex) => processRow(rowIndex, columns, bounds))
  ).then((results) =>
    results.reduce(
      (acc, item) => {
        if (item.status === "fulfilled") acc.succeed.push(item);
        else acc.failed.push(item.reason.message);
        return acc;
      },
      { succeed: [], failed: [] }
    )
  );
}

function computeFilename(options) {
  const format = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Settings")
    .getRange("B3")
    .getValue();
  return format.replace(/\{\{\s*([^\{]+)\s*\}\}/g, (...args) => {
    return options[args[1]];
  });
}

async function processRow(rowIndex, columnsMapping, bounds) {
  try {
    const range = SpreadsheetApp.getActiveSheet().getRange(
      `A${rowIndex}:${indexToColumn(bounds.col)}${rowIndex}`
    );
    const config = getCurrentConfig();
    const blob = DriveApp.getFileById(config.pdfTemplate.id).getBlob();
    const PF = new PdfForm();
    const rowValues = range.getValues()[0];
    const values = rowValues.reduce((acc, value, index) => {
      if (index + 1 in columnsMapping)
        acc.push({
          ref: columnsMapping[index + 1],
          value: value.toString(),
        });
      return acc;
    }, []);
    const data = await PF.setValues(blob, values, true);
    const filenameParts = blob.getName().split(".");
    const extension = filenameParts.pop();
    const newBlob = await PF.saveToPDFBlob({
      data,
      filename:
        computeFilename({
          filename: filenameParts.join("."),
          ...Object.fromEntries(values.map((item) => [item.ref, item.value])),
        }) + `.${extension}`,
    });
    const outputFolder = DriveApp.getFolderById(config.outputFolder.id);
    const newFile = outputFolder.createFile(newBlob);
    const value = SpreadsheetApp.newRichTextValue()
      .setText("OK")
      .setLinkUrl(newFile.getUrl())
      .build();
    range
      .getCell(1, 1)
      .setRichTextValue(value)
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center");
  } catch (error) {
    console.error(error);
    throw new Error(rowIndex);
    SpreadsheetApp.getUi().alert(
      "Une erreur est survenu en générant la ligne " + rowIndex
    );
  }
}

function generatePdf() {
  const ui = SpreadsheetApp.getUi();
  const activeRangeList = SpreadsheetApp.getActiveRangeList();
  const rowsToGenerate = activeRangeList.getRanges().reduce((acc, range) => {
    return acc + range.getNumRows();
  }, 0);

  var confirm = ui.alert(
    `Êtes-vous sûr de vouloir générer ${rowsToGenerate} documents ?`,
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm == ui.Button.OK) {
    ui.alert("Documents générés");
  }
}

/**
 * Displays an HTML-service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API.
 */
function showConfigModal() {
  var html = HtmlService.createHtmlOutputFromFile("Picker.html")
    .setWidth(800)
    .setHeight(650)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, "Select Folder");
}

async function configure(config) {
  const documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty("config", JSON.stringify(config));
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = spreadsheet.getSheetByName("Settings");
  if (!settingsSheet) {
    settingsSheet =
      SpreadsheetApp.getActiveSpreadsheet().insertSheet("Settings");
  }
  settingsSheet
    .getRange("A1:C1")
    .setValues([
      ["Template PDF", config.pdfTemplate.name, config.pdfTemplate.id],
    ]);
  settingsSheet
    .getRange("A2:C2")
    .setValues([
      [
        "Dossier de destination",
        config.outputFolder.name,
        config.outputFolder.id,
      ],
    ]);
  settingsSheet
    .getRange("A3:C3")
    .setValues([
      [
        "Format nom de fichier",
        "",
        "Indiquer un format de nom de fichier dans la cellule précédente, Exemple: {{1378}}_{{filename}} ou 1378 correspond à un ID de champs et filename au nom original",
      ],
    ]);
  settingsSheet
    .getRange("A5:D5")
    .setValues([
      [
        "Afin d'utiliser le générateur, il faut reporter l'ID des champs en première ligne de votre feuille de calcul",
        "",
        "",
        "",
      ],
    ])
    .setFontColor("red")
    .setFontWeight("bold")
    .merge();
  settingsSheet
    .getRange("A7:D7")
    .setValues([
      ["Chargement du PDF et identification des champs...", "", "", ""],
    ])
    .merge();
  const result = await discoverFields(config);
  settingsSheet
    .getRange("A7:D7")
    .setValues([["ID", "Nom", "Type", "Valeur par défaut"]])
    .setBackground("blue")
    .setFontColor("white")
    .setFontWeight("bold");
  result.forEach((field, index) => {
    const rowIndex = 8 + index;
    settingsSheet
      .getRange(`A${rowIndex}:D${rowIndex}`)
      .setValues([[field.ref, field.name, field.type, field.defaultValue]]);
  });
  settingsSheet.setColumnWidth(1, 150);
  settingsSheet.autoResizeColumn(2);
  settingsSheet.setColumnWidths(3, 2, 150);
}

function getCurrentConfig() {
  const documentProperties = PropertiesService.getDocumentProperties();
  const val = documentProperties.getProperty("config");
  return JSON.parse(val);
}

function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}
