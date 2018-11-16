Public Class frmMain

    Private Sub showSCDefaultPropertiesButton_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles showSCDefaultPropertiesButton.Click

        Try
            Dim scriptCtl As MSScriptControl.ScriptControl = New MSScriptControl.ScriptControl()
            Dim stringBldr As System.Text.StringBuilder = New System.Text.StringBuilder()
            With scriptCtl
                stringBldr.Append("AllowUI: ")
                stringBldr.Append(.AllowUI.ToString())
                stringBldr.Append(vbNewLine)
                stringBldr.Append("Language: ")
                If .Language Is Nothing Then
                    .Language = "VBScript"
                    stringBldr.Append(.Language.ToString() & " (was Nothing)")
                Else
                    stringBldr.Append(.Language.ToString())
                End If
                stringBldr.Append(vbNewLine)

                'Notice that the Name property is not accessible when you
                'instantiate a ScriptControl object directly.
                'stringBldr.Append("Name: ")
                'stringBldr.Append(.Name)
                'stringBldr.Append(vbNewLine)

                stringBldr.Append("State: ")
                stringBldr.Append(.State.ToString())
                stringBldr.Append(vbNewLine)
                stringBldr.Append("TimeOut: ")
                stringBldr.Append(.Timeout.ToString())
                stringBldr.Append(vbNewLine)
                stringBldr.Append("UseSafeSubset: ")
                stringBldr.Append(.UseSafeSubset.ToString())
            End With
            MessageBox.Show(stringBldr.ToString(), "Default ScriptControl Object Property Values", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error Ocurred", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub globalModuleSimpleButton_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles globalModuleSimpleButton.Click

        Try
            'One advantage of using an inline direct instantiation of the ScriptControl object
            'is that you don't have to worry about resetting or managing the state of a 
            'shared ScriptControl object that different parts of the code are using.
            'Notice in this procedure how we just create this one and let it go out of
            'scope when we're done.
            Dim scriptCtl As MSScriptControl.ScriptControl = New MSScriptControl.ScriptControl()
            'However, one downside of instantiating your own ScriptControl object is that you
            'have to remember to set the Language property.
            scriptCtl.Language = "VBScript"

            'We're using a simple example here that's hard-coded, but imagine instead
            'that we're loading these scripts from a database or from a file.
            Dim stringBldr As System.Text.StringBuilder = New System.Text.StringBuilder()
            stringBldr.Append("Sub ShowMyName(ByVal firstName, ByVal lastName)")
            stringBldr.Append(vbNewLine)
            stringBldr.Append("    MsgBox(""My name is "" & firstName & "" "" & lastName)")
            stringBldr.Append(vbNewLine)
            stringBldr.Append("End Sub")
            scriptCtl.AddCode(stringBldr.ToString())

            'Notice how we've declared two parameters in the ShowMyName() Sub.
            'We use an array of Object to send to RunCode().
            Dim parms() As Object = {"Super", "Fly"}
            scriptCtl.AllowUI = True
            scriptCtl.Run("ShowMyName", parms)
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error Ocurred", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub showSCompDefaultPropertiesButton_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles showSCompDefaultPropertiesButton.Click

        Try
            Dim stringBldr As System.Text.StringBuilder = New System.Text.StringBuilder()
            'Notice that we're using the object on the form created by dragging the OCX
            'from the Toolbox. Contrast this with showSCDefaultPropertiesButton, which
            'instantiates a ScriptControl object from the COM library. In a simple project,
            'using the OCX approach is suitable and easier, since you don't need to worry
            'about managing the lifetime of your ScriptControl object. On the other hand,
            'if you're using the ScriptControl object in a very narrow scope for a single purpose,
            'it may be cleaner to just instantiate a ScriptControl object inline as in
            'showSCDefaultPropertiesButton.
            With scriptControlOCX
                stringBldr.Append("AllowUI: ")
                stringBldr.Append(.AllowUI.ToString())
                stringBldr.Append(vbNewLine)
                stringBldr.Append("Language: ")
                If .Language Is Nothing Then
                    .Language = "VBScript"
                    stringBldr.Append(.Language.ToString() & " (was Nothing)")
                Else
                    stringBldr.Append(.Language.ToString())
                End If
                stringBldr.Append(vbNewLine)

                stringBldr.Append("Name: ")
                stringBldr.Append(.Name)
                stringBldr.Append(vbNewLine)

                'Notice that the State property is not accessible when the
                'ScriptControl object comes from the OCX.
                'stringBldr.Append("State: ")
                'stringBldr.Append(.State.ToString())
                'stringBldr.Append(vbNewLine)

                stringBldr.Append("TimeOut: ")
                stringBldr.Append(.Timeout.ToString())
                stringBldr.Append(vbNewLine)
                stringBldr.Append("UseSafeSubset: ")
                stringBldr.Append(.UseSafeSubset.ToString())
            End With
            MessageBox.Show(stringBldr.ToString(), "Default ScriptControl Object Property Values", _
                MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error Ocurred", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub modulesExampleButton_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles modulesExampleButton.Click

        Try
            Dim stringBldr As System.Text.StringBuilder = New System.Text.StringBuilder()
            stringBldr.Append("Sub RevealSource")
            stringBldr.Append(vbNewLine)
            stringBldr.Append("    MsgBox ""This message is from the global module.""")
            stringBldr.Append(vbNewLine)
            stringBldr.Append("End Sub")

            Dim parms() As Object = {}
            scriptControlOCX.AllowUI = True
            scriptControlOCX.AddCode(stringBldr.ToString())
            scriptControlOCX.Run("RevealSource", parms)

            stringBldr = new System.Text.StringBuilder()
            stringBldr.Append("Sub RevealSource")
            stringBldr.Append(vbNewLine)
            stringBldr.Append("    MsgBox ""This message is from Module1.""")
            stringBldr.Append(vbNewLine)
            stringBldr.Append("End Sub")

            Dim modOne As MSScriptControl.Module = Nothing
            modOne = scriptControlOCX.Modules.Add("Module1")
            modOne.AddCode(stringBldr.ToString())
            modOne.Run("RevealSource", parms)

        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error Ocurred", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub classExampleButton_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles classExampleButton.Click

        Try
            Dim msgDisplayObject As MessageDisplay = New MessageDisplay()
            'This object will become globally available inside the script module.
            scriptControlOCX.AddObject("msgDisplayObject", msgDisplayObject, True)

            Dim stringBldr As System.Text.StringBuilder = New System.Text.StringBuilder()
            stringBldr.Append("Sub UseDisplayMessage(ByVal messageToShow)")
            stringBldr.Append(vbNewLine)
            stringBldr.Append("    msgDisplayObject.ShowMessage(messageToShow)")
            stringBldr.Append(vbNewLine)
            stringBldr.Append("End Sub")

            Dim parms() As Object = {"This message is very important."}
            scriptControlOCX.AllowUI = True
            scriptControlOCX.AddCode(stringBldr.ToString())
            scriptControlOCX.Run("UseDisplayMessage", parms)
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error Ocurred", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Sub globalModuleNoProcButton_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles globalModuleNoProcButton.Click

        Try
            Dim scriptCtl As MSScriptControl.ScriptControl = New MSScriptControl.ScriptControl()
            scriptCtl.Language = "VBScript"

            Dim stringBldr As System.Text.StringBuilder = New System.Text.StringBuilder()
            stringBldr.Append("MsgBox(""Test"")")

            'Here's the odd behavior: notice how the code we've added is not inside of a
            'Sub or Function; notice how the ScriptControl runs the code right away
            'without calling RunCode(); in fact, there is no way to call RunCode()
            'without a Sub or Function to call.
            scriptCtl.AddCode(stringBldr.ToString())
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error Ocurred", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
End Class
