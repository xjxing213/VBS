Option Explicit

Dim lngOuter
Dim lngInner

For lngOuter = 1 to 5
    MsgBox "Outer loop index: " & lngOuter
    
    For lngInner = 10 to 18 Step 2
        MsgBox "Inner loop index: " & lngInner
    Next
Next
