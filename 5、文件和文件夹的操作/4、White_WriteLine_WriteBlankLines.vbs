'object.CreateTextFile(filename[, overwrite[, unicode]])
'参数：
'object 
'必选项。应为 FileSystemObject 或 Folder 对象的名称。 
'filename 
'必选项。指明所要创建文件的字符串表达式。 
'overwrite 
'可选项。Boolean 值，指明能否覆盖已有文件。如果文件可以覆盖，则值为 true ，否则为 false。如果忽略，则已有文件不能被覆盖。 
'unicode 
'可选项。Boolean 值，指明文件是否以 Unicode 或 ASCII 文件方式创建。如果文件作为 Unicode 文件创建，则值为 true ，如果作为 ASCII 文件创建，则为 false。如果忽略，则假定为 ASCII 文件。 
Dim fso,filepath,f,str

Set fso = CreateObject("sCripting.FileSystemObject")
filepath = "D:\Program Files\Vbsedit\VBSscript\5、文件和文件夹的操作\test.txt"

Set f = fso.CreateTextFile(filePath)

For i = 1 To 100
'将给定的字符串写入到一个 TextStream 文件。
'object.Write(string)

'将指定数量的换行符写入到一个 TextStream 文件。
'object.WriteBlankLines(lines)
'	输入字符后换行
	'方法1
	'不提供换行功能
'	f.Write i
	'相当于写一行回车符
'	f.WriteBlankLines 1
	
	'方法2
	'提供换行
'	f.WriteLine i
	
	'方法3
'	f.Write i & vbCrLf
	
	'方法4 提高效率，减少访问文件次数
	str = str & i & vbCrLf
Next
f.Write str
f.Close

Set fso = Nothing
set f =Nothing

