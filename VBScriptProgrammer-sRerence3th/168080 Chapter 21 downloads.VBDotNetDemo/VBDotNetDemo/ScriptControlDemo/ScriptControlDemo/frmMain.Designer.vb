<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.showSCDefaultPropertiesButton = New System.Windows.Forms.Button
        Me.globalModuleSimpleButton = New System.Windows.Forms.Button
        Me.modulesExampleButton = New System.Windows.Forms.Button
        Me.classExampleButton = New System.Windows.Forms.Button
        Me.scriptControlOCX = New AxMSScriptControl.AxScriptControl
        Me.showSCompDefaultPropertiesButton = New System.Windows.Forms.Button
        Me.globalModuleNoProcButton = New System.Windows.Forms.Button
        Me.tipLabel = New System.Windows.Forms.Label
        CType(Me.scriptControlOCX, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'showSCDefaultPropertiesButton
        '
        Me.showSCDefaultPropertiesButton.Location = New System.Drawing.Point(36, 28)
        Me.showSCDefaultPropertiesButton.Name = "showSCDefaultPropertiesButton"
        Me.showSCDefaultPropertiesButton.Size = New System.Drawing.Size(226, 27)
        Me.showSCDefaultPropertiesButton.TabIndex = 0
        Me.showSCDefaultPropertiesButton.Text = "ScriptControl Object Default Properties"
        Me.showSCDefaultPropertiesButton.UseVisualStyleBackColor = True
        '
        'globalModuleSimpleButton
        '
        Me.globalModuleSimpleButton.Location = New System.Drawing.Point(35, 83)
        Me.globalModuleSimpleButton.Name = "globalModuleSimpleButton"
        Me.globalModuleSimpleButton.Size = New System.Drawing.Size(226, 29)
        Me.globalModuleSimpleButton.TabIndex = 1
        Me.globalModuleSimpleButton.Text = "Simple Code in the Global Module"
        Me.globalModuleSimpleButton.UseVisualStyleBackColor = True
        '
        'modulesExampleButton
        '
        Me.modulesExampleButton.Location = New System.Drawing.Point(36, 143)
        Me.modulesExampleButton.Name = "modulesExampleButton"
        Me.modulesExampleButton.Size = New System.Drawing.Size(226, 29)
        Me.modulesExampleButton.TabIndex = 2
        Me.modulesExampleButton.Text = "Adding Modules and Running Code"
        Me.modulesExampleButton.UseVisualStyleBackColor = True
        '
        'classExampleButton
        '
        Me.classExampleButton.Location = New System.Drawing.Point(35, 201)
        Me.classExampleButton.Name = "classExampleButton"
        Me.classExampleButton.Size = New System.Drawing.Size(226, 29)
        Me.classExampleButton.TabIndex = 3
        Me.classExampleButton.Text = "Exposing Class to Script"
        Me.classExampleButton.UseVisualStyleBackColor = True
        '
        'scriptControlOCX
        '
        Me.scriptControlOCX.Enabled = True
        Me.scriptControlOCX.Location = New System.Drawing.Point(526, 236)
        Me.scriptControlOCX.Name = "scriptControlOCX"
        Me.scriptControlOCX.OcxState = CType(resources.GetObject("scriptControlOCX.OcxState"), System.Windows.Forms.AxHost.State)
        Me.scriptControlOCX.Size = New System.Drawing.Size(38, 38)
        Me.scriptControlOCX.TabIndex = 4
        '
        'showSCompDefaultPropertiesButton
        '
        Me.showSCompDefaultPropertiesButton.Location = New System.Drawing.Point(312, 28)
        Me.showSCompDefaultPropertiesButton.Name = "showSCompDefaultPropertiesButton"
        Me.showSCompDefaultPropertiesButton.Size = New System.Drawing.Size(226, 27)
        Me.showSCompDefaultPropertiesButton.TabIndex = 5
        Me.showSCompDefaultPropertiesButton.Text = "ScriptControl Component Default Properties"
        Me.showSCompDefaultPropertiesButton.UseVisualStyleBackColor = True
        '
        'globalModuleNoProcButton
        '
        Me.globalModuleNoProcButton.Location = New System.Drawing.Point(312, 83)
        Me.globalModuleNoProcButton.Name = "globalModuleNoProcButton"
        Me.globalModuleNoProcButton.Size = New System.Drawing.Size(226, 40)
        Me.globalModuleNoProcButton.TabIndex = 6
        Me.globalModuleNoProcButton.Text = "Interesting Behavior if No Sub or Function Declared"
        Me.globalModuleNoProcButton.UseVisualStyleBackColor = True
        '
        'tipLabel
        '
        Me.tipLabel.Location = New System.Drawing.Point(37, 246)
        Me.tipLabel.Name = "tipLabel"
        Me.tipLabel.Size = New System.Drawing.Size(447, 37)
        Me.tipLabel.TabIndex = 7
        Me.tipLabel.Text = "Tip: to follow along with the code as it executes, run this project from within V" & _
            "isual Studio and set a breakpoint in the code before clicking one of these butto" & _
            "ns."
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(576, 286)
        Me.Controls.Add(Me.tipLabel)
        Me.Controls.Add(Me.globalModuleNoProcButton)
        Me.Controls.Add(Me.showSCompDefaultPropertiesButton)
        Me.Controls.Add(Me.scriptControlOCX)
        Me.Controls.Add(Me.classExampleButton)
        Me.Controls.Add(Me.modulesExampleButton)
        Me.Controls.Add(Me.globalModuleSimpleButton)
        Me.Controls.Add(Me.showSCDefaultPropertiesButton)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VB.NET Script Control Demo - VBScript Programmers Reference, Chapter 18"
        CType(Me.scriptControlOCX, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents showSCDefaultPropertiesButton As System.Windows.Forms.Button
    Friend WithEvents globalModuleSimpleButton As System.Windows.Forms.Button
    Friend WithEvents modulesExampleButton As System.Windows.Forms.Button
    Friend WithEvents classExampleButton As System.Windows.Forms.Button
    Friend WithEvents scriptControlOCX As AxMSScriptControl.AxScriptControl
    Friend WithEvents showSCompDefaultPropertiesButton As System.Windows.Forms.Button
    Friend WithEvents globalModuleNoProcButton As System.Windows.Forms.Button
    Friend WithEvents tipLabel As System.Windows.Forms.Label

End Class
