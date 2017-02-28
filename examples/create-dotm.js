// define global fso
// include JS
includeJS("..\\lib\\json2");
includeJS("..\\lib\\bvba-util");

// read and evaluate JS file
function includeJS(filename) {
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var pathScript = WScript.ScriptFullName;
  var f = fso.GetFile(pathScript);

  filename = fso.GetParentFolderName(f) + "\\" + filename + ".js";
  WScript.Echo(filename);

  var fileStream = fso.openTextFile(filename);
  var fileData = fileStream.readAll();
  fileStream.Close();
  eval(fileData);
}


// iife

(function (config) {

  config = config || {};
  WScript.Echo(JSON.stringify(config));

  var wdFormatXMLTemplateMacroEnabled = 15;






} (BVBA.getConfig()));