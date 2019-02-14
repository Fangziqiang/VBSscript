'函数，函数有返回值，过程没有返回值
'声明 Function 过程的名称、参数以及构成其主体的代码。

' [Public [Default]| Private] Function name [(
'   arglist
')]
'[statements]
'[name = expression]
'[Exit Function] 
'[statements]
'[name = expression]
'End Function 

'参数
'ByVal 
'表示该参数是按值方式传递的。 在函数内部改变该值后，出函数后该值不会改变
'ByRef 
'表示该参数按引用方式传递。在函数内部改变该值后，出函数后该值会改变 
'varname 
'代表参数变量的名称；遵循标准的变量命名规则。 


Dim ary1(3),ary2(3),max
ary1(0) = 1
ary1(1) = 2
ary1(2) = 3

ary2(0) = 4
ary2(1) = 5
ary2(2) = 6

max = GetMax(ary1(0),ary1(1),ary1(2))

MsgBox "最大值是"&"—"&max

Function GetMax(intNum1,intNum2,intNum3)
	If intNum1>intNum2 Then 
		GetMax = intNum1
	Else 
		GetMax = intNum2 
	End If
	'函数的退出
	Exit Function 
	If intNum3>GetMax Then GetMax = intNum3
End Function
