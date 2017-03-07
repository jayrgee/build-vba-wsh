(function () {
  var fso = new ActiveXObject("Scripting.FileSystemObject");

  // iterate command line args
  var objArgs = WScript.Arguments;
  if (objArgs.length > 0) {
    var path = getFullPath(objArgs(0));
    if (path) {
      exportMe(path);
    }
  }

  console.log('bye!');

  function exportMe (filespec) {
    console.log(filespec);

    // spin up app
    var objectName = getObjectNameFromFileExtension(filespec);
    console.log(objectName);

    var app = getAppObject(objectName);

    if (app) {

      exportWord(app, filespec);

      app.Quit();
    }
  }

  function getFullPath(relPath) {
    var filespec = BVBA.getParentFolderName() + "\\" + relPath;

    if (fso.FileExists(filespec)) {
      return filespec;
    }
    return null;
  }

  function getObjectNameFromFileExtension(filespec) {

    var ext = fso.GetExtensionName(filespec);
    ext = ext.toLowerCase();

    if (ext === "doc" || ext === "docm" || ext === "dot" || ext == "dotm") { return "Word.Application"; }

    return null;
  }

  function getAppObject (objectName) {
    try {
      return new ActiveXObject(objectName);

    }
    catch (e) {
      return null;
    }
  }

  function exportWord (app, filespec) {

    var doc = app.Documents.Open(filespec);

    console.log(doc.Name);
    console.log(doc.Type); // wdTypeDocument = 0; wdTypeTemplate = 1

    var exportRoot = BVBA.getParentFolderName(filespec) + "\\export";
    var exportFolder = exportRoot + "\\blah.dotm-" + getTimestamp();

    if (BVBA.checkFolderExists(exportRoot, true)) {
      console.log(BVBA.checkFolderExists(exportFolder, true));
    }

    doc.Close();
  }

  function getTimestamp () {
    var dt = (new Date);
    var mm = dt.getMonth() + 1; // getMonth() is zero-based
    var dd = dt.getDate();
    var hr = dt.getHours();
    var mn = dt.getMinutes();
    var ss = dt.getSeconds();

    return [
      dt.getFullYear(),
      (mm>9 ? '' : '0') + mm,
      (dd>9 ? '' : '0') + dd,
      '-',
      (hr>9 ? '' : '0') + hr,
      (mn>9 ? '' : '0') + mn,
      (ss>9 ? '' : '0') + ss
      ].join('');
  };

  function createExportDir () {

  }
}());
