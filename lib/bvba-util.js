// Create a BVBA object only if one does not already exist. We create the
// methods in a closure to avoid creating global variables.

if (typeof BVBA !== "object") {
    BVBA = {};
}

(function(){
  var fso = new ActiveXObject("Scripting.FileSystemObject");

  if (typeof BVBA.getParentFolderName !== "function") {

    BVBA.getParentFolderName = function getParentFolderName () {
      var pathScript = WScript.ScriptFullName;
      var f = fso.GetFile(pathScript);
      return fso.GetParentFolderName(f);
    };

  }

  if (typeof BVBA.getFileData !== "function") {
    BVBA.getFileData = function getFileData (filepath) {
      var fileStream;
      try
      {
        fileStream = fso.openTextFile(filepath);
      }
      catch(e)
      {
        return null;
      }
      var fileData = fileStream.readAll();
      fileStream.Close();
      return fileData;
    }

    if (typeof BVBA.getConfig !== "function") {
      BVBA.getConfig = function getConfig () {
        var baseName = fso.GetBaseName(WScript.ScriptFullName);
        var filePath = getParentFolderName() + "\\" + baseName + ".json";
        var fileData = getFileData(filePath);
        return JSON.parse(fileData);
      }
    }
  }
}());