Option Explicit

Dim autoRepository,TOCollection,TestObject,i

Set autoRepository = CreateObject("Mercury.ObjectRepositoryUtil")

'���������ļ�
autoRepository.Load "E:\UFTworkstation\Repository2.tsr"

'��ȡ����������Ķ��󼯺�
Set TOCollection = autoRepository.GetAllObjectsByClass("Link")

'�������в��Զ���
For i = 0 To TOCollection.count-1
'��ȡ���Զ���
	Set TestObject = TOCollection.item(i)
	If autoRepository.GetLogicalName(TestObject)="����" Then 
	'���Ķ����� TEXT ����Ϊ��ͼ Ƭ��
		TestObject.SetTOproperty "text","ͼƬ"
	'���¶���
		autoRepository.UpdateObject TestObject
		'��������������
		autoRepository.RenameObject TestObject,"ͼƬ"
		'��������
		autoRepository.Save
	Exit For
	End If
Next
Set testobject = Nothing
Set TOCollection = Nothing
Set autoRepository = Nothing
