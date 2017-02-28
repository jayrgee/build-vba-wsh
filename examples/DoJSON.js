// JScript with JSON!

// define global fso
var fso = new ActiveXObject("Scripting.FileSystemObject");
// and JSON
includeJS("..\\lib\\json2.js");


(function (JSON, fso) {

  var config = getConfig();

  WScript.Echo(JSON.stringify(config));


  function getConfig() {
    var filename = ".\\config.json";
    var fileStream = fso.openTextFile(filename);
    var fileData = fileStream.readAll();
    fileStream.Close();

    return JSON.parse(fileData);
  }

}(this.JSON, this.fso));


function includeJS(filename) {
  var fileStream = fso.openTextFile(filename);
  var fileData = fileStream.readAll();
  fileStream.Close();
  eval(fileData);
}
