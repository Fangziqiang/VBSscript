'WshNetwork ����
Set WshNetwork = WScript.CreateObject("WScript.Network")
WScript.Echo "Domain = " & WshNetwork.UserDomain
WScript.Echo "Computer Name = " & WshNetwork.ComputerName
WScript.Echo "User Name = " & WshNetwork.UserName
