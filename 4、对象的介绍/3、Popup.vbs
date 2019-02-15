'Popup
Dim wsh,desktop,objLnk

Set wsh = CreateObject("WScript.Shell")

MsgBox "我是msgbox"

wsh.Popup "我是msgbox",3

Set wsh = Nothing 