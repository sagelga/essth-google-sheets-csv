<img src="https://img.icons8.com/color/48/000000/google-sheets.png">

# Google Sheets to CSV converter

This script is to convert the Google Sheets data into CSV file, then store the data into Google Drive folder.

Script is licensed as MIT license.

## Parameter

If you liked to use this script on your project, you will have to edit these parameters in order to compatible with your file and folder access.

1. File ID <br>
   This is the file ID that the new CSV will override.<br>
   To retrive the file ID value, use the retrieve all file function.

```js
 try {
    var file = DriveApp.getFileById("1vNB3Wfm919gXdTLbbonpuKaHcOZ__v88");
    file.setContent(csv);
```

2. Folder ID (optional)<br>
   This is the folder that the file will locate. If there is none, the script could work, but when there is no file that is ID as the same as `File ID`, this folder will be the location that the new file will be created. <br>
   Please make sure that there is a valid `File ID` that the script can access.

```js
// catch when there is no file in the folder
var folder = DriveApp.getFolderById("1tTo0PRIkVjbyk5JJoEz-FMzf3JKm8udP");
folder.createFile("File.csv", csv);
```

---

This script is part of Cooperative Studies at King Mongkut Institute of Technology Ladkrabang.
