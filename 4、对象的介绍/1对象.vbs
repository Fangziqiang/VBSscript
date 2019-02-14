'将对象引用赋给一个变量或属性，或者将对象引用与事件关联。

'Set objectvar = {objectexpression | New classname | Nothing}
'-或者-

'Set object.eventname = GetRef(procname)
'创建对象
'CreateObject
'ConnectObject
'GetObject
'对象的属性和方法


Dim wjh

Set wjh = CreateObject("wjh")
'对象的状态有以下几种：
'只读/可写/只读+可写

wjh.xz = "圆"
wjh.qianbi "xiezi"
 

'使用完的对象进行释放
Set wjh = Nothing 
