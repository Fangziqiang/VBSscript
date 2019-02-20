Dim fso 
Set fso = CreateObject("Scripting.FileSystemObject") 
set fn2=fso.GetFile("c:\index2.htm") 

flsize2=fn2.size 
fldate2=fn2.datelastmodified 

set fn=fso.GetFile("c:\index.htm") 

flsize1=fn.size 
fldate1=fn.datelastmodified 
If fso.FileExists("c:\index2.htm") and flsize2>50000 and fldate2>fldate1 Then 
	fso.getfile("c:\index2.htm").copy("c:\index.htm") 
if err.number=0 then WriteHistory "成功"&now(),"log.txt" 
end if 

Sub WriteHistory(hisChars, path) 
Const ForReading = 1, ForAppending = 8 
Dim fso, f 
Set fso = CreateObject("Scripting.FileSystemObject") 
Set f = fso.OpenTextFile(path, ForAppending, True) 
f.WriteLine hisChars 
f.Close 
End Sub


