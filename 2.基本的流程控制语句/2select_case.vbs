' Select Case testexpression
'	[Case expressionlist-n
'		[statements-n]] . . .
'	[Case Else expressionlist-n
'		[elsestatements-n]]
' End Select
'-----------注释：-----------------------
'  如果 testexpression 与任何 Case expressionlist 表达式匹配，则执行此 Case 子句
'  和下一个 Case 子句之间的语句，对于最后的子句，则会执行该子句到 End Select 之
'  间的语句，然后控制权会转到 End Select 之后的语句。如 testexpression 与多个
'  Case 子句中的 expressionlist 表达式匹配，则只有第一个匹配后的语句被执行。
' Case Else 用于指示若在 testexpression 和任何其他 Case 选项的 expressionlist 
' 之间未找到匹配，则执行 elsestatements。虽然不是必要的，但最好是将 Case Else 
' 语句置于 Select Case 块中以处理不可预见的 testexpression 值。如果没有 Case 
' expressionlist 与 testexpression 匹配且无 Case Else 语句，则继续执行 End Select 
' 之后的语句。
'----------------------------------------
Dim Color, MyVar
'改变背景颜色
Sub ChangeBackground (Color)
'lcase(string)仅大写字母转换成小写字母；所有小写字母和非字母字符保持不变。
   MyVar = lcase (Color)
   Select Case MyVar
      Case "red"     document.bgColor = "red"
      Case "green"   document.bgColor = "green"
      Case "blue"    document.bgColor = "blue"
      Case Else      MsgBox "选择另一种颜色"
   End Select
End Sub

ChangeBackground("")

'Select Case测试1
Dim A,B
A=1
B=2
Select Case A+B
      Case "1"     MsgBox "red",,"Select Case测试1"
      Case "2"     MsgBox "green",,"Select Case测试1"
      Case "3"     MsgBox "blue",,"Select Case测试1"
      Case Else      MsgBox "选择另一种颜色",,"Select Case测试1"
End Select

'Select Case测试2
Dim var
var = 1
Select Case True
	Case 1
		MsgBox "var = 1",,"Select Case测试2"
	Case 2
		MsgBox "var = 2",,"Select Case测试2"
	Case 3
		MsgBox "var未知",,"Select Case测试2"
	Case True 
		MsgBox "var = true"
End Select 
'Select Case测试3
Dim var2
var2 = 1
Select Case True
	Case var2 > 0 And var2 < 10
		MsgBox "10 > var >0"
End Select 