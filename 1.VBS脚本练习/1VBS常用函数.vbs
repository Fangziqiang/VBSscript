'VBS�ĳ��ú���
'1������ת������
'CBool(expression)
'expression���㣬�򷵻�false�����򷵻�True
Dim A,B,check
A=5:B=5
check = CBool(A=B)
'True
MsgBox check
A = 0
check = CBool(A)
'False
MsgBox check

'CByte()ת����������Ϊ����
Dim MyDouble,MyByte
MyDouble = 125.5678
MyByte=CByte(MyDouble)
MsgBox MyByte

'Chr()
Dim MyA,MyB
MyA = Chr(65)
MyB = Chr(66)
MsgBox MyA & MyB 

