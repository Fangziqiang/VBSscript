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
 