'FileSystemObject (FSO) ����ģ�ͣ�����Դ��������ԡ��������¼���ʹ�ý���Ϥ��
'object.method �﷨���������ļ��к��ļ���

'ʹ��������ڶ���Ĺ��ߺͣ� 

'HTML ������ Web ҳ 
'Windows Scripting Host ��Ϊ Microsoft Windows �������ļ� 
'Script Control �������������Կ�����Ӧ�ó����ṩ�༭�ű������� 
'���ã�
'FSO ����ģ��ʹӦ�ó����ܴ������ı䡢�ƶ���ɾ���ļ��У���̽���ض����ļ����Ƿ�
'���ڣ������ڣ��������ҳ��й��ļ��е���Ϣ�������ơ������������һ���޸ĵ�����

'����:
'Copy()��CopyFile()��CopyFolder()

'�ƶ�:
'Move()��MoveFile()��MoveFolder()

'������:
'object.Name = [Newname]

Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

'���� �ļ�
fso.CopyFile "D:\FileTest1\�̻�.vbs","D:\FileTest2\"

'fso.CopyFolder "D:\Program Files\Vbsedit","D:\"

'�ƶ��ļ�
If fso.FileExists("D:\FileTest2\vbsedit.key") Then 
	MsgBox "�ļ��Ѵ���"
Else
	fso.MoveFile "D:\FileTest1\vbsedit.key","D:\FileTest2\"
End If

Set fso = Nothing
