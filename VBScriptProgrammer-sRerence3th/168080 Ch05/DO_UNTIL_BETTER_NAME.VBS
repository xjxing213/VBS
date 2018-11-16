Option Explicit

Dim boolStopLooping
Dim lngLoopCount
Dim strResponse

boolStopLooping = True
lngLoopCount = 0
Do
    boolStopLooping = True
    lngLoopCount = lngLoopCount + 1

    strResponse = InputBox("What is the magic word?")
    If UCase(Trim(strResponse)) = "PLEASE" Then
        MsgBox "Correct!  Congratulations!"
    Else
        If lngLoopCount < 5 Then
            MsgBox "Sorry, try again."
            boolStopLooping = False
        Else
            MsgBox "Okay, the word we wanted was 'Please'."
        End If
    End If

Loop Until boolStopLooping = True
