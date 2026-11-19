Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
WshShell.Run "cmd /c for /f ""tokens=5"" %a in ('netstat -aon ^| findstr :3001 ^| findstr LISTENING') do taskkill /f /pid %a", 0, True
WshShell.Run "node server.js", 0, False
WshShell.Run "cmd /c start http://localhost:3001", 0, False
