// define global fso
var fso = new ActiveXObject("Scripting.FileSystemObject");
// and JSON
includeJS(".\\lib\\json2");

// read and evaluate JS file
function includeJS(filename) {
  var fileStream = fso.openTextFile(getParentFolderName() + filename + ".js");
  var fileData = fileStream.readAll();
  fileStream.Close();
  eval(fileData);
}

function getParentFolderName() {
  var pathScript = WScript.ScriptFullName;
  var f = fso.GetFile(pathScript);
  return fso.GetParentFolderName(f);
}

// iterate command line args
objArgs = WScript.Arguments;
for (i = 0; i < objArgs.length; i++) {
  WScript.Echo(objArgs(i));
}

// iife (immediately invoked functional expression)
(function () {

  var config = getConfig();

  WScript.Echo(JSON.stringify(config));


  function getConfig() {
    var filename = ".\\config.json";
    var fileStream;
    try {
      fileStream = fso.openTextFile(filename);
    }
    catch (e) {
      return {};
    }
    var fileData = fileStream.readAll();
    fileStream.Close();
    return JSON.parse(fileData);
  }

}());
