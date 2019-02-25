'如果在 QTP 直接创建 AOM 对象是会报错的，因为 QTP 只允许有一个实例对象，当我们
'已经开启了一个 QTP 之后就不可以再创建另外一个 QTP 了，因此这个时候我们可以直接
'使用 GETOBJECT("","quicktest.application")来对当前启动的 QTP 进行操作,当我们添加下
'图脚本后，我们点击运行后当前 QTP 就会自动隐藏，运行完后自动恢复可见.
'Set qtapp=GetObject("","QuickTest.Application")
'qtapp.visible = false

Dim qtApp

'创建对象
Set qtApp = CreateObject("QuickTest.Application")

'Start QuickTest
qtApp.Launch

'设置为可见
qtApp.Visible = True

'打开测试脚本
qtApp.Open("E:\UFTworkstation\GUITest1")

'运行测试脚本
qtApp.Test.Run

qtApp.Quit
Set qtApp=Nothing
