'以指定次数重复执行一组语句。
' 语法：
'For counter = start To end [Step step]
'[statements]
'[Exit For]
'[statements]
'Next
'counter 
'	用做循环计数器的数值变量。这个变量不能是数组元素或用户自定义类型的元素。 
'start 
'	counter 的初值。 
'end 
'	counter 的终值。 
'step 
'	counter 的步长。如果没有指定，则 step 的默认值为 1。 
'statements 
'	For 和 Next 之间的一条或多条语句，将被执行指定次数。 
' 可以将一个 For...Next 循环放置在另一个 For...Next 循环中，组成嵌套循环。
' 每个循环中的 counter 要使用不同的变量名。
'-------------------------------------
'Test1
Dim a,I,J,K
a=1
For I = 1 To 10
      For J = 1 To 10
            For K = 1 To 10
            a=a+1
            Next
      Next
Next

MsgBox a,,"Test1"

'-------------------------------------
'Test2
Dim b

For b = 0 To 10 Step 1
	MsgBox b,,"Test2" 

Next

'-------------------------------------
'Test3
Dim c

For c = 10 To -6 Step -2 
	MsgBox c,,"Test3" 

Next

'-------------------------------------
'Test4
Dim d

For d = 1 To 10
	MsgBox "d="&d,,"Test4" 
	If d=7 Then Exit For 

Next



