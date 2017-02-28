includeFile("..\\lib\\json2.js");

if (!this.JSON) WScript.Echo("JSON DOESN'T EXIST");

var fso = new ActiveXObject("Scripting.FileSystemObject");
var pathFile = "config.json";

WScript.Echo(pathFile);
var fileStream = fso.openTextFile(pathFile);
var fileData = fileStream.readAll();
fileStream.Close();

var config = JSON.parse(fileData);
WScript.Echo(JSON.stringify(config));




function includeFile(filename) {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var fileStream = fso.openTextFile(filename);
    var fileData = fileStream.readAll();
    fileStream.Close();
    eval(fileData);
}