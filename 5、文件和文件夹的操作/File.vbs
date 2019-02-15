Dim fso,filePath,f1,f2,folderPath

Set fso = CreateObject("Scripting.FileSystemObject")

filepath = "D:\FileTest2\vbsedit.key"
folderPath = "D:\FileTest2"

Set f1 = fso.GetFile(filepath)
Set f2 = fso.GetFolder(folderPath)

f1.Copy  "D:\FileTest1\",False
f2.Name = "FileTest22222"

Set fso = Nothing
Set f1 = Nothing
Set f2 = Nothing

