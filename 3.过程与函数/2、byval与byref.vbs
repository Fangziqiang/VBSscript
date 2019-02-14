'ByVal 
'表示该参数是按值方式传递的。 在函数内部改变该值后，出函数后该值不会改变
'ByRef 
'表示该参数按引用方式传递。在函数内部改变该值后，出函数后该值会改变 
Dim ary1(3),ary2(3),max
ary1(0) = 1
ary1(1) = 2
ary1(2) = 3

ary2(0) = 4
ary2(1) = 5
ary2(2) = 6

GetMax ary1(0),ary1(1),ary1(2)

'输出：1#2000
MsgBox ary1(0) &"#"& ary1(1)

Function GetMax(byval intNum1,byref intNum2,byval intNum3)
	'输出：1#2
	MsgBox intNum1 & "#" & intNum2
	intNum1 = 1000
	intNum2 = 2000
	'输出：1000#2000
	MsgBox intNum1 & "#" & intNum2
End Function
