'���ݱ��ʽ��ֵ��������ִ��һ�����
'If condition Then statements [Else elsestatements ] 

'���ߣ�ʹ�ÿ���ʽ���﷨�� 
'If condition Then
'[statements]
'[ElseIf condition-n Then
'[elseifstatements]] . . .
'[Else
'[elsestatements]]
'End If 

'����1
Dim B
B = 4
A = 3*B
If A>13 Then
	MsgBox "A>13",,"If Then����1"
ElseIf A=10 Then
	MsgBox "A=10",,"If Then����1"
ElseIf A=11 Then
	MsgBox "A=11",,"If Then����1"
ElseIf A=12 Then
	MsgBox "A=12",,"If Then����1"
ElseIf A=13 Then
	MsgBox "A=13",,"If Then����1"
Else
	MsgBox "A<10",,"If Then����1"
End If 

'����2
Dim D
F = 4
E = 3*F
If E>13 Then
	MsgBox "E>13",,"If Then����2"
ElseIf E<13 Then E=E+1
End If 

MsgBox ("E="&E),,"If Then����2"

'����3
Dim intMax,intMin
intMax=1
IntMin=0
'��ֳɶ���д����дEnd If
'If intMax>intMin Then MsgBox intMax & "����" & intMin
If intMax<intMin Then 
	MsgBox intMax & "����" & intMin,,"If Then����3"
Else 
	MsgBox intMax & "С��" & intMin,,"If Then����3"
End If


'����4
Dim intNum
intNum=1

If intNum=0 Then 
	MsgBox intNum & "����0",,"If Then����4"
ElseIf intNum = 1 Then 
	MsgBox intNum & "����1",,"If Then����4"
Else 
	MsgBox intNum & "������0Ҳ������1",,"If Then����4"
End If

'  ��ð�ſ��԰Ѷ��кϲ���һ��
'  _ �»��������е�����
'����5
Dim intNum5
intNum5 = _
1
MsgBox "inuNum5 = " & intNum,,"If Then����5"