'�����������з���ֵ������û�з���ֵ
'���� Function ���̵����ơ������Լ�����������Ĵ��롣

' [Public [Default]| Private] Function name [(
'   arglist
')]
'[statements]
'[name = expression]
'[Exit Function] 
'[statements]
'[name = expression]
'End Function 

'����
'ByVal 
'��ʾ�ò����ǰ�ֵ��ʽ���ݵġ� �ں����ڲ��ı��ֵ�󣬳��������ֵ����ı�
'ByRef 
'��ʾ�ò��������÷�ʽ���ݡ��ں����ڲ��ı��ֵ�󣬳��������ֵ��ı� 
'varname 
'����������������ƣ���ѭ��׼�ı����������� 


Dim ary1(3),ary2(3),max
ary1(0) = 1
ary1(1) = 2
ary1(2) = 3

ary2(0) = 4
ary2(1) = 5
ary2(2) = 6

max = GetMax(ary1(0),ary1(1),ary1(2))

MsgBox "���ֵ��"&"��"&max

Function GetMax(intNum1,intNum2,intNum3)
	If intNum1>intNum2 Then 
		GetMax = intNum1
	Else 
		GetMax = intNum2 
	End If
	'�������˳�
	Exit Function 
	If intNum3>GetMax Then GetMax = intNum3
End Function
