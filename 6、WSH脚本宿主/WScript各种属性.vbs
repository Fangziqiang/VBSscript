'Interactive ����
'���û�ȷ���ű�ģʽ��
'object.Interactive 

'����� JScript ������ʾ�ű�ģʽ��
'WScript.Echo WScript.Interactive;
'����� JScript ���뽫�ű�ģʽ����Ϊ������
'WScript.Interactive = false;

MsgBox WScript.Interactive

'Name ���ԣ�WScript ����
'���� WScript ����������ִ���ļ��������ơ�
'object.Name 
MsgBox WScript.Name

'Path ��ʾ�ű������ļ�Ŀ¼
MsgBox WScript.Path

'ScriptFullName ����,���ص�ǰ���нű�������·��
MsgBox WScript.ScriptFullName

'StdIn ���� (WScript)��ʾ��ǰ�ű���ֻ��������
'Read ����,������������ָ���������ַ���
Dim Input
Input = ""
WScript.Echo "�����룺"
Do While Not WScript.StdIn.AtEndOfLine
   Input = Input & WScript.StdIn.Read(1)
Loop
WScript.Echo "��������ַ���Ϊ"
WScript.Echo Input

'ReadAll ���������������������е��ַ���
'ReadLine ����������������������

'Version ����,���� Windows �ű������İ汾��

MsgBox WScript.Version







