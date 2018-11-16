Dim PartingGreeting
Dim VisitorName

VisitorName = PromptUserName 

If VisitorName <> "" Then
    PartingGreeting = "Goodbye, " & VisitorName & ". Nice to have met you."
 

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
    ElseIf YourName = "abc" Then 
        Greeting = "That's not a real name."	
    ElseIf YourName = "xxx" Then 
        Greeting = "That's not a real name."
    Else
        Greeting = "Hello, " & YourName & ", great to meet you."
    
        If YourName = "Fred" Then
            Greeting = Greeting & "  Nice to see you Fred."
        End If

    End If

    MsgBox Greeting

    PromptUserName = YourName

End Function
