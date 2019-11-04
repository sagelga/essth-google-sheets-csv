/* To view a documentation on how to use each function visit : https://sagelga.github.io/approval-google-addons/ */

var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var sheet = spreadsheet.getActiveSheet();

function onOpen() {
  // Show addon start button
  var submenuEntries = [];

  submenuEntries.push({
    name: "Run",
    functionName: "newFunction"
  });

  spreadsheet.addMenu("CSV", submenuEntries);
}

function newFunction() {
  var csv = "";
  for (var row = 2; row <= sheet.getLastRow(); row += 1) {
    try {
      var rowRange = getRowRange(row, sheet);
      var timeStamp = rowRange[0];

      var rowDate = timeStamp.toDateString();
      var nowDate = new Date().toDateString();

      if (rowDate === nowDate) {
        csv = csv + rowRange + "\r\n";
      }
    } catch (error) {
      console.error("Function error : " + error);
      console.log("row: " + row);
      console.log("rowRange: " + rowRange);
      console.log("rowDate: " + rowDate);
      console.log("nowDate: " + nowDate);

      Browser.msgBox(error);

      throw new Error(error);
    }
  }

  // Check if there is an available storage left in the Drive
  if (DriveApp.getStorageLimit() < 1000) {
    // warn user about low storage size
    console.log("storageLeft: " + DriveApp.getStorageLimit() + ' bytes.');
  }

  try {
    var file = DriveApp.getFileById("1vNB3Wfm919gXdTLbbonpuKaHcOZ__v88");
    file.setContent(csv);
  } catch (error) {
    // catch when there is no file in the folder
    var folder = DriveApp.getFolderById("1tTo0PRIkVjbyk5JJoEz-FMzf3JKm8udP");
    folder.createFile("File.csv", csv);
  }

  console.log("csv: " + csv);
}
