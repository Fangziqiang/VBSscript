'���ѡ��������
'ѡ�����
'	If ... Then ... Else
'	Select Case 
'
'ѭ�����
'	While ... Wend 
'	For ... Next
'	For Each ... Next 
'	Do ... loop

'-----------ѡ�����------------
'1���Ƚ��ٵ�ѡ��ʱ�ʺ�ʹ��If ... Then ... Else
'2�����������ʺ�ʹ��Select Case 
Dim intVar

intVar = InputBox("������һ��������")

Select Case True  
	Case intVar<=100 And intVar >=90
		MsgBox "����" 
	Case intVar<=89 And intVar >=80
		MsgBox "����"  
	Case intVar<=79 And intVar >=70
		MsgBox "��ͨ" 
	Case intVar<=69 And intVar >=60
		MsgBox "����" 
	Case intVar<=59 And intVar >=0
		MsgBox "������" 
	Case Else 
		msgbox "��������"
End Select

'-----------ѭ�����------------
'1��ѭ������ȷ��ʱ�ʺ�ʹ��For Next 
