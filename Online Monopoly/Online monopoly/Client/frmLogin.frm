VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Online Monopoly - Login"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5055
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
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConnect 
      Caption         =   "connect"
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
      Left            =   3480
      Picture         =   "frmLogin.frx":190C8
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtIP1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   13
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox txtIP1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3120
      MaxLength       =   3
      TabIndex        =   12
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox txtIP1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   11
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox txtIP1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   9
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton cmdRegister 
      BackColor       =   &H00C0FFC0&
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
      Left            =   2280
      Picture         =   "frmLogin.frx":32190
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00C0FFC0&
      Caption         =   "login"
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
      Left            =   3600
      Picture         =   "frmLogin.frx":4B258
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0FFC0&
      Caption         =   "exit"
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
      Picture         =   "frmLogin.frx":64320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Timer tmrClientStatus 
      Interval        =   1000
      Left            =   2760
      Top             =   7200
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2280
      Top             =   7200
   End
   Begin VB.Timer tmrConnect 
      Interval        =   1000
      Left            =   1800
      Top             =   7200
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   120
      Picture         =   "frmLogin.frx":7D3E8
      ScaleHeight     =   5145
      ScaleWidth      =   4785
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   4815
      Begin VB.Label lblClientState 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   4920
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   16
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   15
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   14
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IP address : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Username : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   6840
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim loopX As Integer
Dim connectTimeOut As Integer

Private Sub cmdConnect_Click()
    frmMain.wskClient.RemoteHost = txtIP1(0).Text & "." & txtIP1(1).Text & "." & txtIP1(2).Text & "." & txtIP1(3).Text
    frmMain.wskClient.Connect
    For loopX = 0 To 3
        txtIP1(loopX).Enabled = False
    Next
    cmdConnect.Enabled = False
    connectTimeOut = 100
    tmrTimeOut.Enabled = True
    tmrConnect.Enabled = False
End Sub

Private Sub cmdRegister_Click()
    Me.Hide
    frmRegister.Show
End Sub

Private Sub Form_Load()
    If GetSetting("Monopoly", "Game Setting", "Game", 0) = 0 Then
        SaveSetting "Monopoly", "Game Setting", "Game", 1
        SaveSetting "Monopoly", "Game Setting", "Remote Host", "127.0.0.1"
        SaveSetting "Monopoly", "Game Setting", "Remote Port", "5155"
        SaveSetting "Monopoly", "Game Setting", "Sound", True
        SaveSetting "Monopoly", "Game Setting", "Music", True
        SaveSetting "Monopoly", "Game Setting", "Music Volume", 0
        SaveSetting "Monopoly", "Game Setting", "Player Slot Color 1", "&H8080FF"
        SaveSetting "Monopoly", "Game Setting", "Player Slot Color 2", "&HFF8080"
        SaveSetting "Monopoly", "Game Setting", "Player Slot Color 3", "&H80FF80"
        SaveSetting "Monopoly", "Game Setting", "Player Slot Color 4", "&H80FFFF"
    End If
    playSoundBool = GetSetting("Monopoly", "Game Setting", "Sound", True)
    playMusicBool = GetSetting("Monopoly", "Game Setting", "Music", True)
    musicVolume = GetSetting("Monopoly", "Game Setting", "Music Volume", 0)
    frmMain.wskClient.RemoteHost = GetSetting("Monopoly", "Game Setting", "Remote Host", "127.0.0.1")
    frmMain.wskClient.RemotePort = GetSetting("Monopoly", "Game Setting", "Remote Port", "5155")
    For loopX = 0 To 3
        txtIP1(loopX).Text = Split(frmMain.wskClient.RemoteHost, ".")(loopX)
    Next
    LoginEnabled False
End Sub

Private Sub cmdLogin_Click()
    LoginEnabled False
    frmMain.wskClient.SendData "login" & ";" & txtUsername & ";" & txtPassword & ";"
    DoEvents
End Sub

Private Sub tmrConnect_Timer()
    cmdConnect.value = True
End Sub

Private Sub tmrTimeOut_Timer()
    connectTimeOut = connectTimeOut - 1
    If frmMain.wskClient.state <> sckConnecting Then
        If frmMain.wskClient.state <> sckConnected Then
            If frmMain.wskClient.state <> sckClosed Then
                frmMain.wskClient.Close
            End If
            frmMain.wskClient.Connect
        End If
    End If
    If frmMain.wskClient.state = sckConnected Then
        LoginEnabled True
        SaveSetting "Monopoly", "Game Setting", "Remote Host", frmMain.wskClient.RemoteHost
        txtUsername.SetFocus
        connectTimeOut = 0
    End If
    If connectTimeOut <= 0 Then
        tmrTimeOut.Enabled = False
        If frmMain.wskClient.state <> sckConnected Then
            MsgBox "Failed to connect to server", vbExclamation, "Connection"
            If frmMain.wskClient.state <> sckClosed Then
                frmMain.wskClient.Close
                For loopX = 0 To 3
                    txtIP1(loopX).Enabled = True
                Next
                cmdConnect.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Public Sub login(str As String)
    Select Case Split(str, ";")(1)
        Case "Successful":
            PlayerID = Split(Split(str, ";")(2), ",")(0)
            PlayerName = Split(Split(str, ";")(2), ",")(1)
            frmMain.Show
            frmMain.lblPID.Caption = "Player ID : " & PlayerID
            frmMain.lblPName.Caption = "Player Name : " & PlayerName
            Me.Hide
            frmMain.wskClient.SendData "req;pList"
            frmMain.tmrCheckConnection.Enabled = True
            DoEvents
        Case "InvalidPassword":
            MsgBox "Wrong Password", vbExclamation, "Login error"
            txtPassword.Text = ""
            'txtPassword.SetFocus
        Case "InvalidLogin"
            MsgBox "Someone is using this account. Please try again later.", vbExclamation, "Login Error"
        Case "NotExist":
            MsgBox "Username not exist", vbExclamation, "Login error"
            txtPassword.Text = ""
            'txtUsername.SetFocus
    End Select
    LoginEnabled True
    txtPassword.Text = ""
End Sub

Public Sub LoginEnabled(bool As Boolean)
    If bool Then
        txtUsername.BackColor = &HFFFFFF
        txtPassword.BackColor = &HFFFFFF
    Else
        txtUsername.BackColor = &H8000000F
        txtPassword.BackColor = &H8000000F
    End If
    txtUsername.Enabled = bool
    txtPassword.Enabled = bool
    cmdRegister.Enabled = bool
    cmdLogin.Enabled = bool
End Sub

Private Sub tmrClientStatus_Timer()
    lblClientState.Caption = "Client: " & GetStatus(frmMain.wskClient.state)
End Sub

Public Function GetStatus(state As String) As String
    Dim strState As String
    Select Case state
       Case sckClosed
          strState = "Closed"
       Case sckOpen
          strState = "Open"
       Case sckListening
          strState = "Listening"
       Case sckConnectionPending
          strState = "Connection pending"
       Case sckResolvingHost
          strState = "Resolving host"
       Case sckHostResolved
          strState = "Host resolved"
       Case sckConnecting
          strState = "Connecting"
       Case sckConnected
          strState = "Connected"
       Case sckClosing
          strState = "Peer is closing the connection"
       Case sckError
          strState = "Error"
    End Select
    GetStatus = strState
End Function

Private Sub txtIP1_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case Is < 32               ' Control keys are OK.
        Case 48 To 57              ' This is a digit.
        Case Else                  ' Reject any other key.
            KeyAscii = 0
    End Select
End Sub

Private Sub txtIP1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Len(txtIP1(Index).Text) >= 3 And Index <= 2 And KeyCode <> 16 Then
        txtIP1(Index + 1).SetFocus
    End If
End Sub

Private Sub txtIP1_LostFocus(Index As Integer)
    If txtIP1(Index).Text <> "" Then
        If Int(txtIP1(Index).Text) > 255 Then
            txtIP1(Index).Text = 255
        End If
    Else
        txtIP1(Index).Text = 0
    End If
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
