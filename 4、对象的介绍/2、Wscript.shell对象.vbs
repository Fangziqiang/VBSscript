'WshShell对象有三个属性：
'	CurrentDirectory 当前目录
'	Environment 系统环境变量
'	SpecialFolders windows特殊文件夹

'WshShell对象有11个方法：
'	AppActivate
'	CreateShortcut
'	ExpandEnvironmentStrings
'	LogEvent
'	Popup
'	RegRead
'	RegWrite
'	Run 相当于运行
'	SendKeys 模拟键盘操作
'	Exec	运行

'创建快捷方式
Dim wsh,desktop,objLnk

Set wsh = CreateObject("WScript.Shell")
DesktopPath = wsh.SpecialFolders("desktop")
'MsgBox DesktopPath
Set objLnk = wsh.CreateShortcut(DesktopPath & "\testshortcut.lnk")
 