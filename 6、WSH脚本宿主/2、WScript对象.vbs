'WScript 对象---------------------
'提供对 Windows 脚本宿主对象模型根对象的访问。

'说明
'WScript 对象是 Windows 脚本宿主对象模型层次结构的根对象。它从不需要在调用其属性和方法之前进行实例化，并且始终可在任何脚本文件中使用。WScript 对象提供对以下信息的访问： 

'命令行参数， 
'脚本文件的名称， 
'宿主文件名， 
'宿主版本信息。 
'WScript 对象可用来： 

'创建对象， 
'连接对象， 
'与对象断开连接， 
'同步事件， 
'以编程方式停止脚本的执行， 
'将信息输出到默认输出设备（Windows 对话框或命令控制台）。 
'WScript 对象可用来设置脚本运行的模式（交互模式或批处理模式）。

'------属性-----------
'Arguments 属性 | FullName 属性（WScript 对象） | Interactive 属性 | 
'Name 属性 | Path 属性 | ScriptFullName 属性 | ScriptName 属性 | StdErr 属性 | 
'StdIn 属性 | StdOut 属性 | Version 属性

'------方法-----------
'CreateObject 方法 | ConnectObject 方法 | DisconnectObject 方法 | 
'Echo 方法 | GetObject 方法 | Quit 方法 | Sleep 方法


'dir /?  显示dir命令用法

Dim objArgs

Set objArgs = WScript.Arguments

For  i = 0 To objArgs.Count-1
'依次显示WScript.Arguments所有命令
	WScript.Echo "i=" & i & "Item = " & objArgs.Item(i)
Next

