'变量的作用域与存活期
Dim ary(3),max

ary(0) = 0
ary(1) = 1
ary(2) = 2

'获取两个值中的最大值
Private Function GetMun(intNum1,intNum2)
	MsgBox 999
End Function

Function GetMax()
	MsgBox (GetMun (123,45678))
End Function

GetMun 123,45678