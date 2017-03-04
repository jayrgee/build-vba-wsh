// JScript with JSON!

// this script is JScript
// https://en.wikipedia.org/wiki/JScript
// intended to be run by Windows Script Host (CScript.exe)
// https://en.wikipedia.org/wiki/Windows_Script_Host

// JScript is implemented as an Active Scripting engine (like VBScript!) and
// can manipulate Automation objects like "Scripting.FileSystemObject",
// "Word.Application", etc.

// JScript is just another dialect of ECMAScript (aka JavaScript) but the
// version that runs under WSH doesn't support modern features such as JSON.
// However, we can use Scripting.FileSystemObject to load Crockford's JSON
// library https://github.com/douglascrockford/JSON-js



// iterate command line args
objArgs = WScript.Arguments;
for (i = 0; i < objArgs.length; i++) {
  console.log(objArgs(i));
}

// iife (immediately invoked functional expression)
(function () {
  var fso = new ActiveXObject("Scripting.FileSystemObject");

  var config = getConfig();

  console.log(JSON.stringify(config));


  function getConfig() {
    var filespec = getParentFolderName() + "\\config.json";
    console.log(filespec);
    var fileStream;
    try
    {
      fileStream = fso.openTextFile(filespec);
    }
    catch(e)
    {
      return {};
    }
    var fileData = fileStream.readAll();
    fileStream.Close();
    return JSON.parse(fileData);
  }

  function getParentFolderName () {
    var f = fso.GetFile(WScript.ScriptFullName);
    return fso.GetParentFolderName(f);
  }

}());
