Option Explicit

Dim objDict1
Dim objDict2

Set objDict1 = CreateObject("Scripting.Dictionary")
Set objDict2 = CreateObject("Scripting.Dictionary")

objDict1.Add "Hello", "Hello"

MsgBox "The first Dictionary object holds this many items: " & _
    objDict1.Count & vbNewLine & _
    "The second Dictionary object holds this many items: " & _
    objDict2.Count
