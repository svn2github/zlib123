// Create prerequisite subfolder and move installation files there (setup.exe needs it)
var fso = new ActiveXObject("Scripting.FileSystemObject");
if (!fso.FolderExists("KB908002")) {
    fso.CreateFolder("KB908002");
}
fso.MoveFile("extensibilityMSM.msi","KB908002\\");
fso.MoveFile("lockbackRegKey.msi", "KB908002\\");
fso.MoveFile("office2003-kb907417sfxcab-ENU.exe", "KB908002\\");

// Launch setup.exe
var wsh = WScript.CreateObject("WScript.Shell");
wsh.Run("setup.exe", 1, true)