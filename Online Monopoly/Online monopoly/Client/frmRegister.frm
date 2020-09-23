VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Online Monopoly - Register"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Register new user"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Password : "
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Username :"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRegister_Click()
    cmdRegister.Enabled = False
    frmMain.wskClient.SendData "register" & ";" & txtUsername & ";" & txtPassword & ";"
    DoEvents
End Sub

Public Sub register(str As String)
    Select Case Split(str, ";")(1)
        Case "successful":
                Unload Me
        Case "failed":
                MsgBox "Username already exist", vbExclamation + vbOKOnly, "Register"
                txtUsername.SetFocus
                cmdRegister.Enabled = True
                txtPassword.Text = ""
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmLogin.Show
    cmdRegister.Enabled = True
    Unload Me
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Is < 32               ' Control keys are OK.
        Case 48 To 57              ' This is a digit.
        Case 97 To 122
        Case Else                  ' Reject any other key.
            KeyAscii = 0
    End Select
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Is < 32               ' Control keys are OK.
        Case 48 To 57              ' This is a digit.
        Case 97 To 122
        Case Else                  ' Reject any other key.
            KeyAscii = 0
    End Select
End Sub
