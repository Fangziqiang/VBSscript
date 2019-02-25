Option Explicit

Dim autoRepository,TOCollection,TestObject,i

Set autoRepository = CreateObject("Mercury.ObjectRepositoryUtil")

'导入对象库文件
autoRepository.Load "E:\UFTworkstation\Repository2.tsr"

'获取所有链接类的对象集合
Set TOCollection = autoRepository.GetAllObjectsByClass("Link")

'遍历所有测试对象
For i = 0 To TOCollection.count-1
'获取测试对象
	Set TestObject = TOCollection.item(i)
	If autoRepository.GetLogicalName(TestObject)="贴吧" Then 
	'更改对象库的 TEXT 属性为“图 片”
		TestObject.SetTOproperty "text","图片"
	'更新对象
		autoRepository.UpdateObject TestObject
		'重命名对象名称
		autoRepository.RenameObject TestObject,"图片"
		'保存对象库
		autoRepository.Save
	Exit For
	End If
Next
Set testobject = Nothing
Set TOCollection = Nothing
Set autoRepository = Nothing
