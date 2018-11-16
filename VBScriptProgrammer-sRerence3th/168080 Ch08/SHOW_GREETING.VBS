Option Explicit

Dim objGreet
Set objGreet = New Greeting

With objGreet
    .Name = "Dan"
    .ShowGreeting "Informal"
    .ShowGreeting "Formal"
    .ShowGreeting "Casual"
End With
Set objGreet = Nothing

Class Greeting

    Private mstrName

    Public Property Let Name(strName)
        mstrName = strName
    End Property

    Public Sub ShowGreeting(strType)
        MsgBox MakeGreeting(strType) & mstrName & "."
    End Sub

    Private Function MakeGreeting(strType)
        Select Case strType
            Case "Formal"
                MakeGreeting = "Greetings, "
            Case "Informal"
                MakeGreeting = "Hello there, "
            Case "Casual"
                MakeGreeting = "Hey, "
        End Select
    End Function

End Class
