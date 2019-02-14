'当条件为 True 时或条件变为 True 之前重复执行某语句块。

'Do [{While | Until} condition]
'[statements]
'[Exit Do]
'[statements]
'Loop 
'可以使用下面的语法：

'Do
'[Statements]
'[Exit Do]
'[statements]
'Loop [{While | Until} condition]

'Util 后面跟假值时，执行下面语句
'While 后面为真时，执行下面语句

'Test1
Dim i
i = 0
Do Until i>=10 '等价于while i<10
	MsgBox "i = " & i
	i = i+1
Loop

'Test2死循环
' for循环表达式永远为真或者step为0
Dim j

For j=0 To 1 Step 0
	MsgBox "死循环"
Next