'表白记事本
Dim wsh,desktop,objLnk

Set wsh = CreateObject("WScript.Shell")

wsh.Run "notepad"
WScript.Sleep 100

wsh.SendKeys "I"
WScript.Sleep 1000

wsh.SendKeys " "
WScript.Sleep 400

wsh.SendKeys "L"
WScript.Sleep 400

wsh.SendKeys "o"
WScript.Sleep 400

wsh.SendKeys "v"
WScript.Sleep 400

wsh.SendKeys "e"
WScript.Sleep 400

wsh.SendKeys " "
WScript.Sleep 400

wsh.SendKeys "U"
WScript.Sleep 400


Set wsh = Nothing 