Const forReading = 1

Dim fso,filePath,f,i,str

Set fso = CreateObject("Scripting.FileSystemObject")

filePath = "D:\Program Files\Vbsedit\VBSscript\5���ļ����ļ��еĲ���\test.txt"

Set f = fso.OpenTextFile(filePath,forReading)

'��TextStream �ļ��ж�ȡָ���������ַ����������ɴ˵õ����ַ�����
'str = f.Read(6)

'��ȡ TextStream �ļ���ȫ�����ݲ������ɴ˵õ����ַ�����
'str = f.ReadAll

'��TextStream �ļ��ж�ȡһ���У�һֱ�����з��������������з������������ɴ˵õ����ַ�����
str = f.ReadLine



f.Close
MsgBox str
'MsgBox "������"

Set fso = Nothing
Set f = Nothing 

