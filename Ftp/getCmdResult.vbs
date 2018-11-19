Dim strText
Set oShell = CreateObject("WScript.Shell")
Set oExec = oShell.Exec("%COMSPEC% /C ""PING 127.0.0.1""")
Do While Not oExec.StdOut.AtEndOfStream
    strText = oExec.StdOut.ReadAll()
Loop
WScript.Echo strText