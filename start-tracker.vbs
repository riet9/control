Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
psScript = fso.BuildPath(baseDir, "ScreenTimeTracker.ps1")
powerShellPath = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")

extraArgs = ""
For Each arg In WScript.Arguments
    extraArgs = extraArgs & " """ & Replace(arg, """", """""") & """"
Next

command = """" & powerShellPath & """ -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & psScript & """" & extraArgs
shell.Run command, 0, False
