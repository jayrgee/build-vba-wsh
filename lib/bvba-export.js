(function () {
  var fso = new ActiveXObject("Scripting.FileSystemObject");

  // iterate command line args
  var objArgs = WScript.Arguments;
  if (objArgs.length > 0) {
    var path = getFullPath(objArgs(0));
    if (path) {
      exportMe(path);
    }
  }

  console.log('bye!');

  function exportMe (relPath) {
    console.log(relPath);

    // spin up app
    var app = new ActiveXObject("Word.Application");

    console.log('hello!');

    if (app) { app.Quit(); }
  }

  function getFullPath(relPath) {
    var filespec = BVBA.getParentFolderName() + "\\" + relPath;

    if (fso.FileExists(filespec)) {
      return filespec;
    }
    return null;
  }
}());
