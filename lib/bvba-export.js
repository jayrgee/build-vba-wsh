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

    var docName = doc.Name;
    var exportRoot = BVBA.getParentFolderName(filespec) + "\\export";
    var exportFolder = exportRoot + "\\" + docName.replace(/\./g, "_") + "_" + getTimestamp();
    var exportFolderExists;
    var vbaFolder;

    if (BVBA.checkFolderExists(exportRoot, true)) {
      exportFolderExists = BVBA.checkFolderExists(exportFolder, true);
    }

    if (exportFolderExists) {
      vbaFolder = exportFolder + "\\vba";

      if (BVBA.checkFolderExists(vbaFolder, true)) {
        exportVBComponents(doc, vbaFolder)
      };
    }

    doc.Close();
  }

  function exportDocProperties (doc, folderspec) {
    var docName = doc.Name + "json";
    
  }

  function exportVBComponents (doc, folderspec) {

    console.log('exporting vba to ' + folderspec);

    //todo: need to handle error
    // "Programmatic access to Visual Basic Project is not trusted."
    // https://blogs.msdn.microsoft.com/cristib/2012/02/29/vba-how-to-programmatically-enable-access-to-the-vba-object-model-using-macros/

    // ...or set DWORD in registry to value 1:
    // HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Word\Security\AccessVBOM = 1
    forEach(doc.VBProject.VBComponents, function(c) {
      exportVBComponent(c, folderspec);
    });
  }

  function exportVBComponent (c, folderspec) {
    var fileName = c.Name + '.' + getVBExtension(c.Type);
    console.log(' ' + fileName);
    c.Export(folderspec + '\\' + fileName);
  }

  function getVBExtension (vbType) {

    var vbExtensions = {
      1: 'bas',
      2: 'cls',
      3: 'frm',
      11: 'dsr',
      100: 'cls'
    };

    return (vbExtensions[vbType] || 'txt');
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
      (hr>9 ? '' : '0') + hr,
      (mn>9 ? '' : '0') + mn,
      (ss>9 ? '' : '0') + ss
      ].join('');
  };

  function createExportDir () {

  }
  function forEach(collection, func) {
    for (var e = new Enumerator(collection) ; !e.atEnd() ; e.moveNext()) {
      func(e.item());
    }
  }
}());
