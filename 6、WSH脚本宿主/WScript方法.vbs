WScript.Echo "当前时间为：" & Time

WScript.Sleep 3*1000

WScript.Echo "当前时间为：" & Time

'退出不执行后面语句
'WScript.Quit

'CreateShortcut创建桌面快捷方式
Dim WshShell,oShellLink 
set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
set oShellLink = WshShell.CreateShortcut(strDesktop & "\Shortcut Script.lnk")
oShellLink.TargetPath = WScript.ScriptFullName
oShellLink.WindowStyle = 1
'oShellLink.Hotkey = "CTRL+SHIFT+F"
oShellLink.IconLocation = "notepad.exe, 0"
oShellLink.Description = "Shortcut Script"
oShellLink.WorkingDirectory = strDesktop
oShellLink.Save
