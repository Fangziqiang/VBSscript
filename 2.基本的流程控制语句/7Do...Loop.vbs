'������Ϊ True ʱ��������Ϊ True ֮ǰ�ظ�ִ��ĳ���顣

'Do [{While | Until} condition]
'[statements]
'[Exit Do]
'[statements]
'Loop 
'����ʹ��������﷨��

'Do
'[Statements]
'[Exit Do]
'[statements]
'Loop [{While | Until} condition]

'Util �������ֵʱ��ִ���������
'While ����Ϊ��ʱ��ִ���������

'Test1
Dim i
i = 0
Do Until i>=10 '�ȼ���while i<10
	MsgBox "i = " & i
	i = i+1
Loop

'Test2��ѭ��
' forѭ�����ʽ��ԶΪ�����stepΪ0
Dim j

For j=0 To 1 Step 0
	MsgBox "��ѭ��"
Next