'object.CreateTextFile(filename[, overwrite[, unicode]])
'������
'object 
'��ѡ�ӦΪ FileSystemObject �� Folder ��������ơ� 
'filename 
'��ѡ�ָ����Ҫ�����ļ����ַ������ʽ�� 
'overwrite 
'��ѡ�Boolean ֵ��ָ���ܷ񸲸������ļ�������ļ����Ը��ǣ���ֵΪ true ������Ϊ false��������ԣ��������ļ����ܱ����ǡ� 
'unicode 
'��ѡ�Boolean ֵ��ָ���ļ��Ƿ��� Unicode �� ASCII �ļ���ʽ����������ļ���Ϊ Unicode �ļ���������ֵΪ true �������Ϊ ASCII �ļ���������Ϊ false��������ԣ���ٶ�Ϊ ASCII �ļ��� 
Dim fso,filepath,f,str

Set fso = CreateObject("sCripting.FileSystemObject")
filepath = "D:\Program Files\Vbsedit\VBSscript\5���ļ����ļ��еĲ���\test.txt"

Set f = fso.CreateTextFile(filePath)

For i = 1 To 100
'���������ַ���д�뵽һ�� TextStream �ļ���
'object.Write(string)

'��ָ�������Ļ��з�д�뵽һ�� TextStream �ļ���
'object.WriteBlankLines(lines)
'	�����ַ�����
	'����1
	'���ṩ���й���
'	f.Write i
	'�൱��дһ�лس���
'	f.WriteBlankLines 1
	
	'����2
	'�ṩ����
'	f.WriteLine i
	
	'����3
'	f.Write i & vbCrLf
	
	'����4 ���Ч�ʣ����ٷ����ļ�����
	str = str & i & vbCrLf
Next
f.Write str
f.Close

Set fso = Nothing
set f =Nothing

