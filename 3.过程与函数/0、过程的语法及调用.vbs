'���̵��﷨������
Dim ary1(3),ary2(3)
ary1(0) = 1
ary1(1) = 2
ary1(2) = 3

ary2(0) = 4
ary2(1) = 5
ary2(2) = 6

Sub GetMax(intNum1,intNum2,intNum3)
	Dim max
	If intNum1>intNum2 Then 
		max = intNum1
	Else 
		max = intNum2 
	End If
	'���̵��˳�ʹ��Exit Sub ��䣬Exit Sub �������䲻��ִ��
	Exit Sub 
	If intNum3>max Then max = intNum3
	MsgBox "���ֵ��" & max 
End Sub

'���̵ĵ���
'1��ֱ����������
GetMax 1,2,3
'Call GetMax()
Call GetMax(ary1(0),ary1(1),ary1(2))
Call GetMax(ary2(0),ary2(1),ary2(2))