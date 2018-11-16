' Kathie Kingsley-Hughes
' 27 Oct 2003
' This script prompts the user for their name.
' It incorporates various greetings depending on input by the user.
' 
' Added alternative greeting
' Changed variable names to make them more readable 

Dim PartingGreeting
Dim VisitorName

VisitorName = PromptUserName 
'**Start Encode**
If VisitorName <> "" Then
    PartingGreeting = "Hello, " & VisitorName & ". Nice to have met you."
 

Else
    PartingGreeting = "I'm glad to have met you, but I wish I knew your name."
End If

MsgBox PartingGreeting



Function PromptUserName

    ' This Function prompts the user for their name.
    ' It incorporates various greetings depending on input by the user.
    Dim YourName 
    Dim Greeting

    YourName = InputBox("Hello!  What is your name?")

    If YourName = "" Then
        Greeting = "OK.  You don't want to tell me your name."
    Else
        Greeting = "Hello, " & YourName & ", great to meet you."
    End If

    MsgBox Greeting

    PromptUserName = YourName

End Function