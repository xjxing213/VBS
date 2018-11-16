Option Explicit

Private Const DEBUG_MODE = False

If DEBUG_MODE Then
    MsgBox "Script starting."
End If

MsgBox "Non-debug script code executing here."

If DEBUG_MODE Then
    MsgBox "Script ending."
End If
