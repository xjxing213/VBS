Dim WS  
Set WS=CreateObject("MSWinsock.Winsock")  
IPAddress=WS.LocalIP  
MsgBox "Local IP=" & IPAddress