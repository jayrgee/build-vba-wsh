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
  //WScript.Echo(filename);

  var fileStream = fso.openTextFile(filename);
  var fileData = fileStream.readAll();
  fileStream.Close();
  eval(fileData);
}


// iife

(function (config) {

  config = config || {};
  //WScript.Echo(JSON.stringify(config));

  var wdFormatXMLTemplateMacroEnabled = 15;

  var appWd = new ActiveXObject("Word.Application");
  appWd.Visible = true;
  var doc = appWd.Documents.Add();

  var refs = [];
  if (config.VBProject.References) {
    refs = config.VBProject.References;
  }
  var i = 0;
  for (i = 0; i < refs.length; i++) {
    WScript.Echo(refs[i]);
    doc.VBProject.References.AddFromFile(refs[i]);
  }




} (BVBA.getConfig()));