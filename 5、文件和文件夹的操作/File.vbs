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
MsgBox "�ļ���С" & flsize1
'fldate1 = f1.DateLastModified

'fisize2 = f3.Size
'fldate2 = f3.DateLastModified

If fso.FileExists("D:\FileTest1\vbsedit.key") Then 
	MsgBox "�ļ�vbsedit.key�Ѵ���"
'	f1.Copy  "D:\FileTest1\",True
Else 
	f1.Copy  "D:\FileTest1\",True
	MsgBox "�ļ�vbsedit.key���Ƴɹ�"
End If 

f2.Name = "FileTest22222"
If fso.FolderExists("D:\FileTest22222") Then
	MsgBox "�ļ����������޸�ΪFileTest22222"
	f2.Name = "FileTest2"
	If fso.FolderExists ("D:\FileTest2") Then
		MsgBox "�ļ��������������޸�ΪFileTest2"
	Else 
		MsgBox "�Ļ�ԭ��������ʧ��"
	End If
	
Else
	MsgBox "�ļ��������޸�ΪFileTest22222ʱʧ��"
End If 
Set fso = Nothing
Set f1 = Nothing
Set f2 = Nothing

