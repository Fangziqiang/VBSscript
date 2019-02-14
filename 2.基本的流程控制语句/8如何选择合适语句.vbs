'如何选择合适语句
'选择语句
'	If ... Then ... Else
'	Select Case 
'
'循环语句
'	While ... Wend 
'	For ... Next
'	For Each ... Next 
'	Do ... loop

'-----------选择语句------------
'1、比较少的选择时适合使用If ... Then ... Else
'2、数字区间适合使用Select Case 
Dim intVar

intVar = InputBox("请输入一个整数：")

Select Case True  
	Case intVar<=100 And intVar >=90
		MsgBox "优秀" 
	Case intVar<=89 And intVar >=80
		MsgBox "良好"  
	Case intVar<=79 And intVar >=70
		MsgBox "普通" 
	Case intVar<=69 And intVar >=60
		MsgBox "及格" 
	Case intVar<=59 And intVar >=0
		MsgBox "不及格" 
	Case Else 
		msgbox "输入有误"
End Select

'-----------循环语句------------
'1、循环次数确定时适合使用For Next 
