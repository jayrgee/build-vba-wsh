// Create a BVBA object only if one does not already exist. We create the
// methods in a closure to avoid creating global variables.

if (typeof BVBA !== "object") {
  BVBA = {};
}

(function () {
  var fso = new ActiveXObject("Scripting.FileSystemObject");

  BVBA.getApp = function getApp (appId) {
    return new ActiveXObject(appId + '.Application');
  };


  if (typeof BVBA.getParentFolderName !== "function") {

    BVBA.getParentFolderName = function getParentFolderName (filespec) {
      filespec = filespec || WScript.ScriptFullName;
      var f = fso.GetFile(filespec);
      return fso.GetParentFolderName(f);
    };

  }

  if (typeof BVBA.getFileData !== "function") {
    BVBA.getFileData = function getFileData (filepath) {
      var fileStream;
      try {
        fileStream = fso.openTextFile(filepath);
      }
      catch (e) {
        return null;
      }
      var fileData = fileStream.readAll();
      fileStream.Close();
      return fileData;
    }
  }

  if (typeof BVBA.getConfig !== "function") {
    BVBA.getConfig = function getConfig () {
      var baseName = fso.GetBaseName(WScript.ScriptFullName);
      var filePath = getParentFolderName() + "\\" + baseName + ".json";
      var fileData = getFileData(filePath);
      return JSON.parse(fileData);
    }
  }

  if (typeof BVBA.getFilePaths !== "function") {
    BVBA.getFilePaths = function getFilePaths (folderspec) {
      var filePaths = [];
      var fsoFolder;

      if (fso.FolderExists(folderspec)) {
        fsoFolder = fso.GetFolder(folderspec);

        forEach(fsoFolder.Files, function (f) {
          console.log(f.Path);
          filePaths.push(f.Path);
        });
      }
      return filePaths;
    }
  }


  if (typeof BVBA.addReferences !== "function") {
    BVBA.addReferences = function addReferences(doc, refPaths) {
      for (var i = 0; i < refPaths.length; i++) {
        console.log(refPaths[i]);
        doc.VBProject.References.AddFromFile(refPaths[i]);
      }
    }
  }


  if (typeof BVBA.importVBAComponents !== "function") {
    BVBA.importVBAComponents = function importVBAComponents (doc, vbaPaths) {
      // import VBA components to VBProject
      for (var i = 0; i < vbaPaths.length; i++) {
        importVbaComponent(doc, vbaPaths[i]);
      }
    }
  }

  function importVbaComponent (doc, vbaPath) {
    if (isVbaComponentExtensionValid(fso.GetExtensionName(vbaPath))) {
      console.log(vbaPath);
      doc.VBProject.VBComponents.Import(vbaPath);
    }
  }

  function isVbaComponentExtensionValid (extension) {
    var ext = extension.toLowerCase();
    return (ext === "bas" || ext === "cls" || ext === "frm") ? true : false;
  }


  if (typeof BVBA.saveWordDocument !== "function") {
    BVBA.saveWordDocument = function saveWordDocument (doc, folderspec, name, extension) {

      var saveFormat = getWdFormatFromExtension(extension);
      var docPath = folderspec + "\\" + name + "." + extension;
      console.log(docPath);

      if (checkFolderExists(folderspec, true)) {
        doc.SaveAs(docPath, saveFormat);
      }
    }
  }

  if (typeof BVBA.checkFolderExists !== "function") {
    BVBA.checkFolderExists = checkFolderExists;
  }

  // Check folder exists
  function checkFolderExists (folderspec, doCreate) {
    if (doCreate && doCreate === true) {
      if (!fso.FolderExists(folderspec)) {
        fso.CreateFolder(folderspec);
      }
    }
    return fso.FolderExists(folderspec);
  }

  // Get Word Format
  function getWdFormatFromExtension (extension) {
    var wdFormatDocument97 = 0; // .doc
    var wdFormatTemplate97 = 1; // .dot
    var wdFormatXMLDocumentMacroEnabled = 13; // .docm
    var wdFormatXMLTemplateMacroEnabled = 15; // .dotm

    if (extension.toLowerCase() === "doc") { return wdFormatDocument97; }
    if (extension.toLowerCase() === "dot") { return wdFormatTemplate97; }
    if (extension.toLowerCase() === "docm") { return wdFormatXMLDocumentMacroEnabled; }
    if (extension.toLowerCase() === "dotm") { return wdFormatXMLTemplateMacroEnabled; }
    return null;
  }

  // Enumerate items in collection
  // The JScript Enumerator object provides a way to access any member of a
  // collection and behaves similarly to the For...Each statement in VBScript.
  function forEach(collection, func) {
    for (var e = new Enumerator(collection) ; !e.atEnd() ; e.moveNext()) {
      func(e.item());
    }
  }

}());
