'FileSystemObject (FSO) 对象模型，允许对大量的属性、方法和事件，使用较熟悉的
'object.method 语法，来处理文件夹和文件。

'使用这个基于对象的工具和： 

'HTML 来创建 Web 页 
'Windows Scripting Host 来为 Microsoft Windows 创建批文件 
'Script Control 来对用其他语言开发的应用程序提供编辑脚本的能力 
'作用：
'FSO 对象模型使应用程序能创建、改变、移动和删除文件夹，或探测特定的文件夹是否
'存在，若存在，还可以找出有关文件夹的信息，如名称、被创建或最后一次修改的日期

'复制:
'Copy()、CopyFile()、CopyFolder()

'移动:
'Move()、MoveFile()、MoveFolder()

'重命名:
'object.Name = [Newname]

Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

'复制 文件
fso.CopyFile "D:\FileTest1\绿化.vbs","D:\FileTest2\"

'fso.CopyFolder "D:\Program Files\Vbsedit","D:\"

'移动文件
If fso.FileExists("D:\FileTest2\vbsedit.key") Then 
	MsgBox "文件已存在"
Else
	fso.MoveFile "D:\FileTest1\vbsedit.key","D:\FileTest2\"
End If

Set fso = Nothing
