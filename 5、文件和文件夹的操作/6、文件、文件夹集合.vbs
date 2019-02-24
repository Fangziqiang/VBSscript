Const ForReading = 1
Dim fso,filePath,f,fc,f1,str

Set fso  = CreateObject("Scripting.FileSystemObject")

Set f = fso.GetFolder("D:\")

'目录下的文件
Set fc = f.Files

'子文件夹对象
Set fc2 = f.SubFolders

'从一个对象中获取子对象
For Each f1 In fc
	MsgBox f1.name
Next

For Each f2 In fc2
	MsgBox f2.name
Next

Set fso = Nothing
Set fc = Nothing
Set f = Nothing



