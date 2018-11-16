VERSION 5.00
Begin VB.Form frmConnection 
   Caption         =   "Connection String"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5640
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboCombo 
      Height          =   315
      Left            =   1680
      TabIndex        =   22
      Text            =   "Combo"
      Top             =   2640
      Width           =   3855
   End
   Begin VB.TextBox txtKey 
      BackColor       =   &H80000011&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   21
      Text            =   "Key"
      Top             =   2040
      Width           =   3855
   End
   Begin VB.TextBox txtSubpath 
      BackColor       =   &H80000011&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   19
      Text            =   "Subpath"
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   14
      Text            =   "Text"
      Top             =   4770
      Width           =   3855
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   10
      Text            =   "Text"
      Top             =   4410
      Width           =   3855
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   9
      Text            =   "Text"
      Top             =   4050
      Width           =   3855
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Text            =   "Text"
      Top             =   3690
      Width           =   3855
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   7
      Text            =   "Text"
      Top             =   3330
      Width           =   3855
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text"
      Top             =   2970
      Width           =   3855
   End
   Begin VB.Label lblKey 
      Alignment       =   1  'Right Justify
      Caption         =   "Key :"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblSubpath 
      Alignment       =   1  'Right Justify
      Caption         =   "Subpath :"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblCombo 
      Alignment       =   1  'Right Justify
      Caption         =   "Connection Type:"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblRegistry 
      Caption         =   "Registry"
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label lblRegistryValue 
      Alignment       =   1  'Right Justify
      Caption         =   "Registry Key Value:"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Label"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Label"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Label"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Label"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Label"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Label"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblExplanation 
      Caption         =   "Explanation"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents objScript As MSScriptControl.ScriptControl
Attribute objScript.VB_VarHelpID = -1

Private Sub cboCombo_Click()
    objScript.Run "cboCombo_Click"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdProcess_Click()
    objScript.Run "cmdProcess_Click"
End Sub

Private Sub cmdRegister_Click()
    RegistrySave txtSubpath.Text, txtKey.Text, lblRegistry.Caption
    MsgBox "Registration Complete"
    cmdExit.Visible = True
End Sub

Private Sub Form_Load()
    Set objScript = InitScriptControl(Me)
    objScript.Run "init"
End Sub

Private Sub txtText_Change(Index As Integer)
    On Error Resume Next
    objScript.Run "txtText_Change", Index
End Sub

Private Sub txtText_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    KeyAscii = objScript.Run("txtText_KeyPress", Index, KeyAscii)
End Sub
