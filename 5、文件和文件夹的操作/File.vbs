Option Explicit

Dim fso,filePath,f1,f2,folderPath,filepath1,filepath2,flsize1

Set fso = CreateObject("Scripting.FileSystemObject")

filepath1 = "D:\FileTest2\vbsedit.key"
filepath2 = "D:\FileTest1\vbsedit.key"
folderPath = "D:\FileTest2"

Set f1 = fso.GetFile(filepath1)
'Set f3 = fso.GetFile(filepath2)
Set f2 = fso.GetFolder(folderPath)

flsize1 = f1.Size
MsgBox "文件大小" & flsize1
'fldate1 = f1.DateLastModified

'fisize2 = f3.Size
'fldate2 = f3.DateLastModified

If fso.FileExists("D:\FileTest1\vbsedit.key") Then 
	MsgBox "文件vbsedit.key已存在"
'	f1.Copy  "D:\FileTest1\",True
Else 
	f1.Copy  "D:\FileTest1\",True
	MsgBox "文件vbsedit.key复制成功"
End If 

f2.Name = "FileTest22222"
If fso.FolderExists("D:\FileTest22222") Then
	MsgBox "文件夹名称已修改为FileTest22222"
	f2.Name = "FileTest2"
	If fso.FolderExists ("D:\FileTest2") Then
		MsgBox "文件夹名称已重新修改为FileTest2"
	Else 
		MsgBox "改回原来的名称失败"
	End If
	
Else
	MsgBox "文件夹名称修改为FileTest22222时失败"
End If 
Set fso = Nothing
Set f1 = Nothing
Set f2 = Nothing

