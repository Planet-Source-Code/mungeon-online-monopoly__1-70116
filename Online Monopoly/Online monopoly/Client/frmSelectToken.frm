VERSION 5.00
Begin VB.Form frmSelectToken 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Monopoly"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9825
   ControlBox      =   0   'False
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
   Picture         =   "frmSelectToken.frx":0000
   ScaleHeight     =   7965
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00FFC0C0&
      Caption         =   "select"
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
      Left            =   4560
      Picture         =   "frmSelectToken.frx":190C8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Personality"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   5895
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
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
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   5655
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please choose your token."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5895
   End
   Begin VB.Image imgToken 
      Height          =   975
      Index           =   0
      Left            =   7440
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   975
   End
End
Attribute VB_Name = "frmSelectToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loopX As Integer

Private Sub cmdSelect_Click()
    frmMain.wskClient.SendData "cmd;changeToken;" & tableID & ";" & PlayerNumber & ";" & selectedToken
    DoEvents
    Unload Me
    Monopoly.SetFocus
End Sub

Private Sub Form_Load()
    Dim intTop As Long
    Dim intLeft As Long
    Set activeWindow = Me
    intTop = 400
    intLeft = 200
    For loopX = 1 To 10
        Load imgToken(loopX)
        imgToken(loopX).Picture = LoadPicture("images/Token/" & Token(loopX).file)
        imgToken(loopX).Visible = True
        imgToken(loopX).Top = intTop
        imgToken(loopX).Left = intLeft
        intLeft = intLeft + 1200
        If loopX = 5 Then
            intTop = intTop + 1200
            intLeft = 200
        End If
    Next
    imgToken(0).Visible = False
    cmdSelect.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set activeWindow = Monopoly
End Sub

Private Sub imgToken_Click(Index As Integer)
    If Index >= 1 And Index <= 10 Then
        For loopX = 1 To 10
            imgToken(loopX).BorderStyle = 0
        Next
        imgToken(Index).BorderStyle = 1
        selectedToken = Index
        lblName.Caption = Token(Index).name
        lblDescription.Caption = Token(Index).description
        If Not cmdSelect.Enabled Then
            cmdSelect.Enabled = True
        End If
    End If
End Sub
