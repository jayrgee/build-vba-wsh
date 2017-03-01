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
  //var fso = new ActiveXObject("Scripting.FileSystemObject");

  config = config || {};
  //WScript.Echo(JSON.stringify(config));

  // spin up app
  var appWd = new ActiveXObject("Word.Application");

  // create a new doc
  var doc = appWd.Documents.Add();


  // add VBProject references
  var refs = [];
  if (config.VBProject.References) {
    refs = config.VBProject.References;
  }
  BVBA.addReferences(doc, refs);


  // get paths of VBA components to be imported
  var vbaPaths = [];
  var vbaRootPath;

  if (config.VBProject.VBSource) {
    vbaRootPath = BVBA.getParentFolderName() + "\\" + config.VBProject.VBSource;
    WScript.Echo(vbaRootPath);
    vbaPaths = BVBA.getFilePaths(vbaRootPath);
  }
  
  // import VBA components to VBProject
  BVBA.importVBAComponents(doc, vbaPaths);

  // Save and close doc
  var wdFormatXMLTemplateMacroEnabled = 15;
  var docName;
  var docExtension;
  if (config.Document) {
    docName = config.Document.Name || "New Document";
    docExtension = config.Document.Extension || "docm";
  }
  var docPath = BVBA.getParentFolderName() + "\\" + docName + "." + docExtension;
  WScript.Echo(docPath);
  doc.SaveAs(docPath, wdFormatXMLTemplateMacroEnabled);
  doc.Close();

  // Quit app
  appWd.Quit();

  WScript.Echo("Done!")

} (BVBA.getConfig()));