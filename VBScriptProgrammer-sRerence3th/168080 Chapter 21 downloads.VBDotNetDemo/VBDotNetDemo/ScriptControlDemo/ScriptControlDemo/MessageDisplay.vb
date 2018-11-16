Public Class MessageDisplay
    Public Sub ShowMessage(ByVal message As String)
        MessageBox.Show("This message is coming from an instance of the MessageDisplay class: " & message, _
            "MessageDisplay Class Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class
