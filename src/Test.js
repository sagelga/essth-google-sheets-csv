function tester() {
  var files = DriveApp.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    Logger.log(file.getName());
    Logger.log(file.getId());
  }
}
