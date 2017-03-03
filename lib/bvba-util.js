// Create a BVBA object only if one does not already exist. We create the
// methods in a closure to avoid creating global variables.

if (typeof BVBA !== "object") {
  BVBA = {};
}

(function () {
  var fso = new ActiveXObject("Scripting.FileSystemObject");

  BVBA.getApp = function getApp(appId) {
    return new ActiveXObject(appId + '.Application');
  };


  if (typeof BVBA.getParentFolderName !== "function") {

    BVBA.getParentFolderName = function getParentFolderName() {
      var pathScript = WScript.ScriptFullName;
      var f = fso.GetFile(pathScript);
      return fso.GetParentFolderName(f);
    };

  }

  if (typeof BVBA.getFileData !== "function") {
    BVBA.getFileData = function getFileData(filepath) {
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
    BVBA.getConfig = function getConfig() {
      var baseName = fso.GetBaseName(WScript.ScriptFullName);
      var filePath = getParentFolderName() + "\\" + baseName + ".json";
      var fileData = getFileData(filePath);
      return JSON.parse(fileData);
    }
  }

  if (typeof BVBA.getFilePaths !== "function") {
    BVBA.getFilePaths = function getFilePaths(folderspec) {
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
    BVBA.importVBAComponents = function importVBAComponents(doc, vbaPaths) {
      // import VBA components to VBProject
      var vbaPath;
      var xName;
      for (var i = 0; i < vbaPaths.length; i++) {
        vbaPath = vbaPaths[i];
        xName = fso.GetExtensionName(vbaPath).toLowerCase();
        if (xName === "bas" || xName === "cls" || xName === "frm") {
          console.log(vbaPath);
          doc.VBProject.VBComponents.Import(vbaPath)
        }
      }
    }
  }


  if (typeof BVBA.saveWordDocument !== "function") {
    BVBA.saveWordDocument = function saveWordDocument(doc, path, name, extension) {
      var wdFormatXMLDocumentMacroEnabled = 13;
      var wdFormatXMLTemplateMacroEnabled = 15;

      var saveFormat = wdFormatXMLDocumentMacroEnabled;
      if (extension.toLowerCase() === "dotm") {
        saveFormat = wdFormatXMLTemplateMacroEnabled;
      }
      var docPath = path + "\\" + name + "." + extension;
      console.log(docPath);
      doc.SaveAs(docPath, saveFormat);
    }
  }


  // Enumerate items in collection
  function forEach(collection, func) {
    for (var e = new Enumerator(collection) ; !e.atEnd() ; e.moveNext()) {
      func(e.item());
    }
  }

}());