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

    var doc;

    if (app) {

      doc = app.Documents.Open(filespec);

      console.log(doc.Name);
      console.log(doc.Type); // wdTypeDocument = 0; wdTypeTemplate = 1

      doc.Close();



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

}());
