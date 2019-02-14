'当指定的条件为 True 时，执行一系列的语句。

' While condition
' &nbsp；  Version [statements]
' Wend

'Test1
Dim i
i=0

While i<10
	i = i+1
	MsgBox "i=" & i,1,"Test1"
wend