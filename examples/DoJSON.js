var fso = new ActiveXObject("Scripting.FileSystemObject");
includeJS("..\\lib\\json2.js");


(function (JSON, fso) {
  var pathFile = "config.json";

  WScript.Echo(pathFile);
  var fileStream = fso.openTextFile(pathFile);
  var fileData = fileStream.readAll();
  fileStream.Close();

  var config = JSON.parse(fileData);
  WScript.Echo(JSON.stringify(config));

}(this.JSON, fso));


function includeJS(filename) {
  var fileStream = fso.openTextFile(filename);
  var fileData = fileStream.readAll();
  fileStream.Close();
  eval(fileData);
}