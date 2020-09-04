// Mapeamento de unidade de rede
var objNet = WScript.CreateObject("WScript.Network");
var objFso = WScript.CreateObject("Scripting.FileSystemObject");
var path = "\\\\saont46\\apps4";

if (!objFso.DriveExists("T:"))
// objNet.RemoveNetworkDrive("T:", true, true);
objNet.MapNetworkDrive("T:", path, false);​