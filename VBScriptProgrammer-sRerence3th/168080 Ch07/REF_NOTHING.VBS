Option Explicit

Dim objDict1
Dim objDict2

Set objDict1 = CreateObject("Scripting.Dictionary")
'The object now exists and has a reference count of 1
Set objDict2 = objDict1
'The object now has a reference count of 2

objDict1.Add "Hello", "Hello"

MsgBox "The first Dictionary object holds this many " & _
    " items: " & objDict1.Count & vbNewLine & _
    "The second Dictionary object holds this many " & _
    "items: " & objDict2.Count

Set objDict1 = Nothing
'The object still exists because objDict2 still
'holds a reference
Set objDict2 = Nothing
'The object's reference count has now gone down to 0,
'so it has been destroyed.
