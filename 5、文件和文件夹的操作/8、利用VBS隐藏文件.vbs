Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const Hidden =2
Const ReadOnly = 1

Dim fso,f,str

Set fso  = CreateObject("Scripting.FileSystemObject")

Set f = fso.GetFile("D:\Program Files\Vbsedit\VBSscript\5���ļ����ļ��еĲ���\test.txt")

'���ļ�������Ϊֻ��������
f.Attributes = Hidden + ReadOnly

MsgBox "�������"

Set fso = Nothing
Set f = Nothing



