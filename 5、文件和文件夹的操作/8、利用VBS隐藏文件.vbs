Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const Hidden =2
Const ReadOnly = 1

Dim fso,f,str

Set fso  = CreateObject("Scripting.FileSystemObject")

Set f = fso.GetFile("D:\Program Files\Vbsedit\VBSscript\5、文件和文件夹的操作\test.txt")

'将文件属性设为只读和隐藏
f.Attributes = Hidden + ReadOnly

MsgBox "操作完成"

Set fso = Nothing
Set f = Nothing



