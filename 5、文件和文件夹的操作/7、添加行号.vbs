Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim fso,f,str

Set fso  = CreateObject("Scripting.FileSystemObject")

Set f = fso.OpenTextFile("D:\Program Files\Vbsedit\VBSscript\5���ļ����ļ��еĲ���\test.txt",ForReading)

'AtEndOfLine�ж��ļ��Ƿ񵽴�ĩβ
While f.AtEndOfLine <> True
	i = i+1
	str = str & "��"& i & "�У�" _
		& f.ReadLine & vbCrLf
Wend
f.Close

Set f = fso.OpenTextFile("D:\Program Files\Vbsedit\VBSscript\5���ļ����ļ��еĲ���\test.txt",ForWriting)
f.Write str
f.Close
MsgBox "�������"

Set f = fso.OpenTextFile("D:\Program Files\Vbsedit\VBSscript\5���ļ����ļ��еĲ���\test.txt",ForAppending)
f.Write str
f.Close
MsgBox "׷��д��������"

Set fso = Nothing
Set f = Nothing



