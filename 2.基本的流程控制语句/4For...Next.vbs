'��ָ�������ظ�ִ��һ����䡣
' �﷨��
'For counter = start To end [Step step]
'[statements]
'[Exit For]
'[statements]
'Next
'counter 
'	����ѭ������������ֵ�����������������������Ԫ�ػ��û��Զ������͵�Ԫ�ء� 
'start 
'	counter �ĳ�ֵ�� 
'end 
'	counter ����ֵ�� 
'step 
'	counter �Ĳ��������û��ָ������ step ��Ĭ��ֵΪ 1�� 
'statements 
'	For �� Next ֮���һ���������䣬����ִ��ָ�������� 
' ���Խ�һ�� For...Next ѭ����������һ�� For...Next ѭ���У����Ƕ��ѭ����
' ÿ��ѭ���е� counter Ҫʹ�ò�ͬ�ı�������
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



