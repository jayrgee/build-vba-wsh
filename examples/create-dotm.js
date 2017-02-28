// define global fso
var fso = new ActiveXObject("Scripting.FileSystemObject");
// and JSON
includeJS("..\\lib\\json2");

// read and evaluate JS file
function includeJS(filename) {
  filename = getParentFolderName() + "\\" + filename + ".js";
  WScript.Echo(filename);

  var fileStream = fso.openTextFile(filename);
  var fileData = fileStream.readAll();
  fileStream.Close();
  eval(fileData);
}

function getParentFolderName() {
  var pathScript = WScript.ScriptFullName;
  var f = fso.GetFile(pathScript);
  return fso.GetParentFolderName(f);
}

function getFileData(filepath) {
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

function getConfig() {
  var baseName = fso.GetBaseName(WScript.ScriptFullName);
  var filePath = getParentFolderName() + "\\" + baseName + ".json";
  var fileData = getFileData(filePath);
  return JSON.parse(fileData);
}


// iife

(function (config) {

  config = config || {};
  WScript.Echo(JSON.stringify(config));

  var wdFormatXMLTemplateMacroEnabled = 15;






} (getConfig()));