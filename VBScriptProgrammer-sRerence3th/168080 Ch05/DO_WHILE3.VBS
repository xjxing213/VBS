Option Explicit

Dim boolLoopAgain
Dim lngLoopCount
Dim strResponse

boolLoopAgain = False
lngLoopCount = 0
Do
    boolLoopAgain = False
    lngLoopCount = lngLoopCount + 1

    strResponse = InputBox("What is the magic word?")
    If UCase(Trim(strResponse)) = "PLEASE" Then
        MsgBox "Correct!  Congratulations!"
    Else
        If lngLoopCount < 5 Then
            MsgBox "Sorry, try again."
            boolLoopAgain = True
        Else
            MsgBox "Okay, the word we wanted was 'Please'."
            Exit Do
        End If
    End If

Loop While boolLoopAgain
