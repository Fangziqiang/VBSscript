'WScript ����---------------------
'�ṩ�� Windows �ű���������ģ�͸�����ķ��ʡ�

'˵��
'WScript ������ Windows �ű���������ģ�Ͳ�νṹ�ĸ��������Ӳ���Ҫ�ڵ��������Ժͷ���֮ǰ����ʵ����������ʼ�տ����κνű��ļ���ʹ�á�WScript �����ṩ��������Ϣ�ķ��ʣ� 

'�����в����� 
'�ű��ļ������ƣ� 
'�����ļ����� 
'�����汾��Ϣ�� 
'WScript ����������� 

'�������� 
'���Ӷ��� 
'�����Ͽ����ӣ� 
'ͬ���¼��� 
'�Ա�̷�ʽֹͣ�ű���ִ�У� 
'����Ϣ�����Ĭ������豸��Windows �Ի�����������̨���� 
'WScript ������������ýű����е�ģʽ������ģʽ��������ģʽ����

'------����-----------
'Arguments ���� | FullName ���ԣ�WScript ���� | Interactive ���� | 
'Name ���� | Path ���� | ScriptFullName ���� | ScriptName ���� | StdErr ���� | 
'StdIn ���� | StdOut ���� | Version ����

'------����-----------
'CreateObject ���� | ConnectObject ���� | DisconnectObject ���� | 
'Echo ���� | GetObject ���� | Quit ���� | Sleep ����


'dir /?  ��ʾdir�����÷�

'Arguments ���ԣ�WScript ����
'���� WshArguments ���󣨲���������
'object.Arguments

Dim objArgs,sum,s1,s2

'1��Arguments����
Set objArgs = WScript.Arguments

' dir /a /b

'For  i = 0 To objArgs.Count-1
'������ʾWScript.Arguments��������,������������ʾ���ýű�ִ�з�ʽ��
'���������У�cscript WScript����.vbs dir /a /b
'	WScript.Echo "i= " & i & " Item = " & objArgs.Item(i)
	
'Next

'1+��nΪֹ�Ľ��
s1 = objArgs.Item(0)
s2 = objArgs.Item(1)
sum = 0
For i = s1 To s2
	sum = sum +i
Next

'cscript WScript����.vbs 1 100 �����1�ӵ�100�ĺ�
WScript.Echo "��" & s1 & "��" & s2 _
	& "֮�������֮��Ϊ" & sum