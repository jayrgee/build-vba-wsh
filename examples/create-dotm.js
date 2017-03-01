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
  var fso = new ActiveXObject("Scripting.FileSystemObject");

  config = config || {};
  //WScript.Echo(JSON.stringify(config));

  var wdFormatXMLTemplateMacroEnabled = 15;

  // spin up app
  var appWd = new ActiveXObject("Word.Application");

  // create a new doc
  var doc = appWd.Documents.Add();

  var i = 0;

  // add VBProject references
  var refs = [];
  if (config.VBProject.References) {
    refs = config.VBProject.References;
  }

  for (i = 0; i < refs.length; i++) {
    WScript.Echo(refs[i]);
    doc.VBProject.References.AddFromFile(refs[i]);
  }

  // get paths of VBA components to be imported
  var vbaPaths = [];
  var vbaRootPath;
  var oVBAFolder;

  if (config.VBProject.VBSource) {
    vbaRootPath = BVBA.getParentFolderName() + "\\" + config.VBProject.VBSource;
    WScript.Echo(vbaRootPath);
    if (fso.FolderExists(vbaRootPath)) {
      oVBAFolder = fso.GetFolder(vbaRootPath);

      forEach(oVBAFolder.Files, function (f) {
        WScript.Echo(f.Path);
        vbaPaths.push(f.Path);
      });
    }
  }
  
  // import VBA components to VBProject
  var vbaPath;
  var xName;
  for (i = 0; i < vbaPaths.length; i++) {
    vbaPath = vbaPaths[i];
    xName = fso.GetExtensionName(vbaPath).toLowerCase();
    if (xName === "bas" || xName === "cls" || xName === "frm" ) {
      WScript.Echo(vbaPath);
      doc.VBProject.VBComponents.Import(vbaPath)
    }
  }

  // Save and close doc
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

  // Enumerate items in collection
  function forEach(collection, func) {
    for (var e = new Enumerator(collection) ; !e.atEnd() ; e.moveNext()) {
      func(e.item());
    }
  }

} (BVBA.getConfig()));