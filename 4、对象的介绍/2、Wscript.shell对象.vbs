'WshShell�������������ԣ�
'	CurrentDirectory ��ǰĿ¼
'	Environment ϵͳ��������
'	SpecialFolders windows�����ļ���

'WshShell������11��������
'	AppActivate
'	CreateShortcut
'	ExpandEnvironmentStrings
'	LogEvent
'	Popup
'	RegRead
'	RegWrite 
'	Run �൱������
'	SendKeys ģ����̲���
'	Exec	����

'������ݷ�ʽ
Dim wsh,desktop,objLnk

Set wsh = CreateObject("WScript.Shell")
DesktopPath = wsh.SpecialFolders("desktop")
'MsgBox DesktopPath
Set objLnk = wsh.CreateShortcut(DesktopPath & "\testshortcut.lnk")
'objLnk.Arguments = "1 2 3"
objLnk.Description = "test shortcut"
'objLnk.IconLocation = "vbsedit.exe,3"
objLnk.WindowStyle = 3
objLnk.TargetPath = "D:\Program Files\Vbsedit\vbsedit.exe"
objLnk.WorkingDirectory = "D:\Program Files\Vbsedit"
objLnk.Save

Set wsh = Nothing
Set objLnk = Nothing
