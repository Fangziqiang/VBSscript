' Select Case testexpression
'	[Case expressionlist-n
'		[statements-n]] . . .
'	[Case Else expressionlist-n
'		[elsestatements-n]]
' End Select
'-----------ע�ͣ�-----------------------
'  ��� testexpression ���κ� Case expressionlist ���ʽƥ�䣬��ִ�д� Case �Ӿ�
'  ����һ�� Case �Ӿ�֮�����䣬���������Ӿ䣬���ִ�и��Ӿ䵽 End Select ֮
'  �����䣬Ȼ�����Ȩ��ת�� End Select ֮�����䡣�� testexpression ����
'  Case �Ӿ��е� expressionlist ���ʽƥ�䣬��ֻ�е�һ��ƥ������䱻ִ�С�
' Case Else ����ָʾ���� testexpression ���κ����� Case ѡ��� expressionlist 
' ֮��δ�ҵ�ƥ�䣬��ִ�� elsestatements����Ȼ���Ǳ�Ҫ�ģ�������ǽ� Case Else 
' ������� Select Case �����Դ�����Ԥ���� testexpression ֵ�����û�� Case 
' expressionlist �� testexpression ƥ������ Case Else ��䣬�����ִ�� End Select 
' ֮�����䡣
'----------------------------------------
Dim Color, MyVar
'�ı䱳����ɫ
Sub ChangeBackground (Color)
'lcase(string)����д��ĸת����Сд��ĸ������Сд��ĸ�ͷ���ĸ�ַ����ֲ��䡣
   MyVar = lcase (Color)
   Select Case MyVar
      Case "red"     document.bgColor = "red"
      Case "green"   document.bgColor = "green"
      Case "blue"    document.bgColor = "blue"
      Case Else      MsgBox "ѡ����һ����ɫ"
   End Select
End Sub

ChangeBackground("")

'Select Case����1
Dim A,B
A=1
B=2
Select Case A+B
      Case "1"     MsgBox "red",,"Select Case����1"
      Case "2"     MsgBox "green",,"Select Case����1"
      Case "3"     MsgBox "blue",,"Select Case����1"
      Case Else      MsgBox "ѡ����һ����ɫ",,"Select Case����1"
End Select

'Select Case����2
Dim var
var = 1
Select Case True
	Case 1
		MsgBox "var = 1",,"Select Case����2"
	Case 2
		MsgBox "var = 2",,"Select Case����2"
	Case 3
		MsgBox "varδ֪",,"Select Case����2"
	Case True 
		MsgBox "var = true"
End Select 
'Select Case����3
Dim var2
var2 = 1
Select Case True
	Case var2 > 0 And var2 < 10
		MsgBox "10 > var >0"
End Select 