Const forReading = 1

Dim fso,filePath,f,i,str

Set fso = CreateObject("Scripting.FileSystemObject")

filePath = "D:\Program Files\Vbsedit\VBSscript\5、文件和文件夹的操作\test.txt"

Set f = fso.OpenTextFile(filePath,forReading)

'从TextStream 文件中读取指定数量的字符，并返回由此得到的字符串。
'str = f.Read(6)

'读取 TextStream 文件的全部内容并返回由此得到的字符串。
'str = f.ReadAll

'从TextStream 文件中读取一整行（一直到换行符，但不包括换行符），并返回由此得到的字符串。
str = f.ReadLine



f.Close
MsgBox str
'MsgBox "已运行"

Set fso = Nothing
Set f = Nothing 

