VERSION 5.00
Begin VB.Form frmCreate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Online Monopoly - Create New Game"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4590
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
   Palette         =   "frmCreate.frx":0000
   PaletteMode     =   2  'Custom
   Picture         =   "frmCreate.frx":190C8
   ScaleHeight     =   1935
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboNumOfPlayer 
      Appearance      =   0  'Flat
      Height          =   360
      ItemData        =   "frmCreate.frx":32190
      Left            =   2160
      List            =   "frmCreate.frx":3219D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFC0&
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
      Left            =   240
      Picture         =   "frmCreate.frx":321AA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      BackColor       =   &H00C0FFC0&
      Caption         =   "create"
      Default         =   -1  'True
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
      Left            =   2880
      Picture         =   "frmCreate.frx":4B272
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   4440
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   4440
      X2              =   120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Player:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Set activeWindow = Me
    cboNumOfPlayer.Text = 4
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()
    Dim temp As Integer
    Dim num1 As Integer
    Dim num2 As Integer
    Dim i As Integer
    If txtTitle.Text = "" Then
        MsgBox "Please enter game title", vbExclamation + vbOKOnly, "Create Game"
        txtTitle.SetFocus
    Else
        gameTitle = txtTitle.Text
        maxPlayer = cboNumOfPlayer.Text
        currentRules = 0
        frmMain.wskClient.SendData "table;create;" & gameTitle & ";" & maxPlayer & ";" & currentRules & ";"
        DoEvents
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set activeWindow = frmMain
End Sub


Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Is < 32               ' Control keys are OK.
        Case 48 To 57              ' This is a digit.
        Case 97 To 122
        Case Else                  ' Reject any other key.
            KeyAscii = 0
    End Select
End Sub
