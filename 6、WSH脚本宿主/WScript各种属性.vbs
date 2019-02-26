'Interactive 属性
'设置或确定脚本模式。
'object.Interactive 

'下面的 JScript 代码显示脚本模式。
'WScript.Echo WScript.Interactive;
'下面的 JScript 代码将脚本模式设置为批处理。
'WScript.Interactive = false;

MsgBox WScript.Interactive

'Name 属性（WScript 对象）
'返回 WScript 对象（主机可执行文件）的名称。
'object.Name 
MsgBox WScript.Name

'Path 显示脚本宿主文件目录
MsgBox WScript.Path

'ScriptFullName 属性,返回当前运行脚本的完整路径
MsgBox WScript.ScriptFullName

'StdIn 属性 (WScript)显示当前脚本的只读输入流
'Read 方法,从输入流返回指定数量的字符。
Dim Input
Input = ""
WScript.Echo "请输入："
Do While Not WScript.StdIn.AtEndOfLine
   Input = Input & WScript.StdIn.Read(1)
Loop
WScript.Echo "你输入的字符串为"
WScript.Echo Input

'ReadAll 方法，从输入流返回所有的字符。
'ReadLine 方法，从输入流返回整行

'Version 属性,返回 Windows 脚本宿主的版本。

MsgBox WScript.Version







