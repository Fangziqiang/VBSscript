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

Dim objArgs

Set objArgs = WScript.Arguments

For  i = 0 To objArgs.Count-1
'������ʾWScript.Arguments��������
	WScript.Echo "i=" & i & "Item = " & objArgs.Item(i)
Next

