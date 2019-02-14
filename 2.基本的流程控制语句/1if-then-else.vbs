'根据表达式的值有条件地执行一组语句
'If condition Then statements [Else elsestatements ] 

'或者，使用块形式的语法： 
'If condition Then
'[statements]
'[ElseIf condition-n Then
'[elseifstatements]] . . .
'[Else
'[elsestatements]]
'End If 

'测试1
Dim B
B = 4
A = 3*B
If A>13 Then
	MsgBox "A>13",,"If Then测试1"
ElseIf A=10 Then
	MsgBox "A=10",,"If Then测试1"
ElseIf A=11 Then
	MsgBox "A=11",,"If Then测试1"
ElseIf A=12 Then
	MsgBox "A=12",,"If Then测试1"
ElseIf A=13 Then
	MsgBox "A=13",,"If Then测试1"
Else
	MsgBox "A<10",,"If Then测试1"
End If 

'测试2
Dim D
F = 4
E = 3*F
If E>13 Then
	MsgBox "E>13",,"If Then测试2"
ElseIf E<13 Then E=E+1
End If 

MsgBox ("E="&E),,"If Then测试2"

'测试3
Dim intMax,intMin
intMax=1
IntMin=0
'拆分成多行写必须写End If
'If intMax>intMin Then MsgBox intMax & "大于" & intMin
If intMax<intMin Then 
	MsgBox intMax & "大于" & intMin,,"If Then测试3"
Else 
	MsgBox intMax & "小于" & intMin,,"If Then测试3"
End If


'测试4
Dim intNum
intNum=1

If intNum=0 Then 
	MsgBox intNum & "等于0",,"If Then测试4"
ElseIf intNum = 1 Then 
	MsgBox intNum & "等于1",,"If Then测试4"
Else 
	MsgBox intNum & "不等于0也不等于1",,"If Then测试4"
End If

'  ：冒号可以把多行合并成一行
'  _ 下划线起到续行的作用
'测试5
Dim intNum5
intNum5 = _
1
MsgBox "inuNum5 = " & intNum,,"If Then测试5"