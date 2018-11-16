Option Explicit

Dim objDict1
Dim objDict2

Set objDict1 = CreateObject("Scripting.Dictionary")
Set objDict2 = objDict1

objDict1.Add "Hello", "Hello"

MsgBox "The first Dictionary object holds this many " & _
    "items: " & objDict1.Count & vbNewLine & _
    "The second Dictionary object holds this many " & _
    "items: " & objDict2.Count
