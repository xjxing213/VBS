Option Explicit

Dim astrColors(3)
Dim strElement

astrColors(0) = "Red"
astrColors(1) = "Green"
astrColors(2) = "Blue"
astrColors(3) = "Yellow"

For Each strElement In astrColors
    MsgBox strElement
Next
