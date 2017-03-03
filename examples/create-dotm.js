(function () {
  var config = BVBA.getConfig() || {};

  //WScript.Echo(JSON.stringify(config));

  // spin up app
  var appWd = new ActiveXObject("Word.Application");

  // create a new doc
  var doc = appWd.Documents.Add();

  // Set VBProject Name
  if (config.VBProject.Name) {
    doc.VBProject.Name = config.VBProject.Name;
  }

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
  var docName;
  var docExtension;
  if (config.Document) {
    docName = config.Document.Name || "New Document";
    docExtension = config.Document.Extension || "docm";
  }
  BVBA.saveWordDocument(doc, BVBA.getParentFolderName(), docName, docExtension);
  doc.Close();

  // Quit app
  appWd.Quit();

  WScript.Echo("Done!")

}());