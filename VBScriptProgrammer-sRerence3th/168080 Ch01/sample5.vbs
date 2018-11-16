Dim Greeting
Dim YourName
Dim TryAgain

Do
    TryAgain = "No"

    YourName = InputBox("Please enter your name:")

    If YourName = "" Then
        MsgBox "You must enter your name to continue."
        TryAgain = "Yes"
    Else
        Greeting = "Hello, " & YourName & ", great to meet you."
    End If
  
Loop While TryAgain = "Yes"

MsgBox Greeting
