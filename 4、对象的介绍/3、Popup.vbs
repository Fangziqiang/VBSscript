'Popup
Dim wsh,desktop,objLnk

Set wsh = CreateObject("WScript.Shell")

MsgBox "����msgbox"

wsh.Popup "����msgbox",3

Set wsh = Nothing 