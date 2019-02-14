'过程的语法及调用
Dim ary1(3),ary2(3)
ary1(0) = 1
ary1(1) = 2
ary1(2) = 3

ary2(0) = 4
ary2(1) = 5
ary2(2) = 6

Sub GetMax(intNum1,intNum2,intNum3)
	Dim max
	If intNum1>intNum2 Then 
		max = intNum1
	Else 
		max = intNum2 
	End If
	'过程的退出使用Exit Sub 语句，Exit Sub 后面的语句不会执行
	Exit Sub 
	If intNum3>max Then max = intNum3
	MsgBox "最大值是" & max 
End Sub

'过程的调用
'1、直接输入名字
GetMax 1,2,3
'Call GetMax()
Call GetMax(ary1(0),ary1(1),ary1(2))
Call GetMax(ary2(0),ary2(1),ary2(2))