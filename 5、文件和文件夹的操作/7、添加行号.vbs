Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim fso,f,str

Set fso  = CreateObject("Scripting.FileSystemObject")

Set f = fso.OpenTextFile("D:\Program Files\Vbsedit\VBSscript\5、文件和文件夹的操作\test.txt",ForReading)

'AtEndOfLine判断文件是否到达末尾
While f.AtEndOfLine <> True
	i = i+1
	str = str & "第"& i & "行：" _
		& f.ReadLine & vbCrLf
Wend
f.Close

Set f = fso.OpenTextFile("D:\Program Files\Vbsedit\VBSscript\5、文件和文件夹的操作\test.txt",ForWriting)
f.Write str
f.Close
MsgBox "操作完成"

Set f = fso.OpenTextFile("D:\Program Files\Vbsedit\VBSscript\5、文件和文件夹的操作\test.txt",ForAppending)
f.Write str
f.Close
MsgBox "追加写入操作完成"

Set fso = Nothing
Set f = Nothing



