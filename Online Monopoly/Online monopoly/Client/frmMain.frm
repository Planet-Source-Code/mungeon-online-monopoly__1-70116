VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Online Monopoly - Main"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11055
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
   ScaleHeight     =   8295
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSetting 
      Caption         =   "setting"
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
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Timer tmrMusic 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   0
      Top             =   960
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrCheckConnection 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   480
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   10815
      Begin VB.Label lblPID 
         BackStyle       =   0  'Transparent
         Caption         =   "player ID:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label lblPName 
         BackStyle       =   0  'Transparent
         Caption         =   "player name: "
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   4095
      End
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      Height          =   2535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4560
      Width           =   8535
   End
   Begin MSComctlLib.ListView lvwPlayer 
      Height          =   3615
      Left            =   8880
      TabIndex        =   4
      Top             =   4560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   3563
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "send"
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
      Left            =   7320
      Picture         =   "frmMain.frx":190C8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      MaxLength       =   80
      TabIndex        =   5
      Top             =   7200
      Width           =   7095
   End
   Begin MSComctlLib.ListView lvwGameTable 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Rules"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Max Player"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Player 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Player 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Player 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Player 4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "join"
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
      Left            =   6720
      Picture         =   "frmMain.frx":32190
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "create"
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
      Left            =   4800
      Picture         =   "frmMain.frx":4B258
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   120
      Picture         =   "frmMain.frx":64320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Image imgBG 
      Height          =   8310
      Index           =   1
      Left            =   8280
      Picture         =   "frmMain.frx":7D3E8
      Top             =   0
      Width           =   8310
   End
   Begin VB.Image imgBG 
      Height          =   8310
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":964B0
      Top             =   0
      Width           =   8310
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim loopX As Integer
Dim strTemp As String
                
Private Sub cmdSetting_Click()
    Load frmSetting
    frmSetting.Show
End Sub

Private Sub Form_Activate()
    activeWindow.Show
    activeWindow.SetFocus
End Sub

Private Sub Form_Load()
    Set activeWindow = Me
    tmrMusic.Enabled = True
    Call rulesSetting
    Call loadToken
End Sub

Private Sub cmdExit_Click()
    If MsgBox("Are you sure you wan to exit game?", vbQuestion + vbYesNo, "Monopoly") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseMidi
    LogOut
    End
End Sub

Public Sub LogOut()
    If wskClient.state = sckConnected Then
        wskClient.SendData "logout;"
        DoEvents
        wskClient.Close
    End If
End Sub

Private Sub tmrCheckConnection_Timer()
    If wskClient.state <> sckConnected Then
        MsgBox "You have been disconnected from server.", vbExclamation + vbOKOnly, "Disconnected"
        Unload Me
    End If
End Sub

Private Sub tmrMusic_Timer()
    On Error Resume Next
    If playMusicBool Then
        If v_dmss.GetSeek >= v_dms.GetLength Then
            CloseMidi
            PlayMidi (1)
        End If
    End If
End Sub

'###### Chatting ##### Begin
Private Sub txtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSend.value = True
    End If
End Sub

Private Sub cmdSend_Click()
    If Trim$(txtSend) <> "" Then
        wskClient.SendData "msg;public;" & txtSend
        DoEvents
        txtSend.Text = ""
        txtSend.SetFocus
    End If
End Sub
'###### Chatting ##### End

Private Sub cmdCreate_Click()
    frmCreate.Show
End Sub

Private Sub cmdJoin_Click()
    Dim countPlayer As Integer
    If lvwGameTable.ListItems.count > 0 Then
        If lvwGameTable.ListItems(lvwGameTable.SelectedItem.Index).ListSubItems(8).Text = "Waiting" Then
            countPlayer = 0
            For loopX = 4 To 7
                If lvwGameTable.ListItems(lvwGameTable.SelectedItem.Index).ListSubItems(loopX).Text <> "" Then
                    countPlayer = countPlayer + 1
                End If
            Next
            If countPlayer < Int(lvwGameTable.ListItems(lvwGameTable.SelectedItem.Index).ListSubItems(3).Text) Then
                wskClient.SendData "table;reqGameInfo;" & lvwGameTable.ListItems(lvwGameTable.SelectedItem.Index).Text & ";"
                DoEvents
            Else
                MsgBox "The table are fulled. Please choose others table.", vbExclamation + vbOKOnly, "Join Game"
            End If
        End If
    End If
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
    Dim strmsg As String
    Dim z As Integer
    Dim a As Integer
    wskClient.GetData strmsg
    If strmsg <> "" Then
        Select Case Split(strmsg, ";")(0)
            Case "register"
                frmRegister.register strmsg
            Case "login"
                frmLogin.login strmsg
            Case "tpList" '##### add/remove data from player and table list
                strmsg = Right(strmsg, Len(strmsg) - Len(Split(strmsg, ";")(0)) - 1)
                addRemoveToList strmsg
            Case "updateTList"
                On Error Resume Next
                lvwGameTable.ListItems(Int(Split(strmsg, ";")(1))).ListSubItems(Int(Split(strmsg, ";")(2)) + 3).Text = Split(strmsg, ";")(4)
            Case "msg" '##### add message to chat box
                txtMsg.Text = txtMsg.Text & vbCrLf & Right(strmsg, Len(strmsg) - (Len(Split(strmsg, ";")(0)) + 1))
                txtMsg.SelStart = Len(txtMsg.Text)
            Case "pList"
                strmsg = Right(strmsg, Len(strmsg) - Len(Split(strmsg, ";")(0)) - 1)
                listPlayer strmsg
            Case "tList"
                strmsg = Right(strmsg, Len(strmsg) - Len(Split(strmsg, ";")(0)) - 1)
                listTable strmsg
            Case "setTableID"
                tableID = Split(strmsg, ";")(1)
                PlayerNumber = 1
                player(PlayerNumber).tokenID = 1
                player(1).keyPlayer = True
                Unload frmCreate
                Unload Monopoly
                frmMain.Hide
                Load Monopoly
                Monopoly.Show
            Case "reqGameInfo"
                strTemp = tableID & "," & gameTitle & "," & maxPlayer & "," & currentRules
                For z = 1 To 4
                    strTemp = strTemp & "," & player(z).PID & "," & player(z).PName & "," & player(z).tokenID & "," & player(z).ready & "," & player(z).keyPlayer
                Next
                strTemp = Replace(strTemp, "False", 0)
                strTemp = Replace(strTemp, "true", 1)
                wskClient.SendData "table;gameInfo;" & Split(strmsg, ";")(1) & ";" & strTemp & ";"
                DoEvents
            Case "receiveGameInfo"
                Dim i As Integer
                strTemp = Split(strmsg, ";")(1)
                tableID = Split(strTemp, ",")(0)
                gameTitle = Split(strTemp, ",")(1)
                maxPlayer = Split(strTemp, ",")(2)
                currentRules = Split(strTemp, ",")(3)
                i = 4
                For z = 1 To 4
                    player(z).PID = Split(strTemp, ",")(i)
                    player(z).PName = Split(strTemp, ",")(i + 1)
                    player(z).tokenID = Split(strTemp, ",")(i + 2)
                    player(z).ready = Split(strTemp, ",")(i + 3)
                    player(z).keyPlayer = Split(strTemp, ",")(i + 4)
                    i = i + 5
                Next
                wskClient.SendData "table;join;" & tableID & ";" & PlayerID & ";" & PlayerName
                DoEvents
            Case "playerJoin"
                If Int(PlayerID) = Int(Split(strmsg, ";")(2)) Then
                    PlayerNumber = Split(strmsg, ";")(1)
                    player(PlayerNumber).tokenID = 1
                    Unload Monopoly
                    frmMain.Hide
                    Load Monopoly
                    Monopoly.Show
                Else
                    player(Int(Split(strmsg, ";")(1))).PID = Split(strmsg, ";")(2)
                    player(Int(Split(strmsg, ";")(1))).PName = Split(strmsg, ";")(3)
                    player(Int(Split(strmsg, ";")(1))).tokenID = 1
                    Monopoly.lblWPJoinName(Int(Split(strmsg, ";")(1))).Caption = Split(strmsg, ";")(3)
                    Monopoly.imgWPToken(Int(Split(strmsg, ";")(1))).Picture = LoadPicture(App.Path & "/images/token/" & token(player(Int(Split(strmsg, ";")(1))).tokenID).file)
                End If
            Case "playerQuit"
                Call Monopoly.playerquit(Int(Split(strmsg, ";")(1)))
            Case "tableMsg"
                Call Monopoly.addChatMsg(strmsg)
            Case "startGame"
                Call Monopoly.initGameVar(strmsg)
            Case "updateStartedGame"
                lvwGameTable.ListItems(Int(Split(strmsg, ";")(1))).ListSubItems(8).Text = "Playing"
            Case "changeToken"
                Call Monopoly.changeToken(Int(Split(strmsg, ";")(1)), Int(Split(strmsg, ";")(2)))
            Case "changeRules"
                Call Monopoly.changeRules(Int(Split(strmsg, ";")(1)), Int(Split(strmsg, ";")(2)))
            Case "playerReady"
                Call Monopoly.playerReady(Int(Split(strmsg, ";")(1)))
            Case "payTax"
                Call Monopoly.payTax(Int(Split(strmsg, ";")(1)))
            Case "payPercentTax"
                Call Monopoly.payPercentTax(Int(Split(strmsg, ";")(1)))
            Case "payFine"
                Call Monopoly.payJailFine(Int(Split(strmsg, ";")(1)))
            Case "useCard"
                Call Monopoly.useCard(Int(Split(strmsg, ";")(1)))
            Case "rollDiceResult"
                Call Monopoly.rollDiceResult(Int(Split(strmsg, ";")(1)), Int(Split(strmsg, ";")(2)), Int(Split(strmsg, ";")(3)))
            Case "buyProperties"
                Call Monopoly.buyProperties(Int(Split(strmsg, ";")(1)), Int(Split(strmsg, ";")(2)))
            Case "proposal"
                Call Monopoly.receiveProposal(strmsg)
            Case "acceptTrade"
                Call Monopoly.acceptTrade(strmsg)
            Case "rejectTrade"
                Call Monopoly.rejectTrade(strmsg)
            Case "auction"
                Call Monopoly.auction(Int(Split(strmsg, ";")(1)), Int(Split(strmsg, ";")(2)))
            Case "auctionBid"
                Call Monopoly.addBidAmount(Int(Split(strmsg, ";")(1)), Int(Split(strmsg, ";")(2)))
            Case "mortgage"
                Call Monopoly.mortgage(Int(Split(strmsg, ";")(1)))
            Case "mortgageDeedCard"
                Call Monopoly.mortgageDeedCard(Int(Split(strmsg, ";")(1)), Int(Split(strmsg, ";")(2)))
            Case "unmortgageDeedCard"
                Call Monopoly.unmortgageDeedCard(Int(Split(strmsg, ";")(1)), Int(Split(strmsg, ";")(2)))
            Case "mortgageClose"
                Call Monopoly.closeMortgage
            Case "buildHouse"
                Call Monopoly.buildHouse(Int(Split(strmsg, ";")(1)), Int(Split(strmsg, ";")(2)))
            Case "sellHouse"
                Call Monopoly.sellHouse(Int(Split(strmsg, ";")(1)), Int(Split(strmsg, ";")(2)))
            Case "done"
                Call Monopoly.done(Int(Split(strmsg, ";")(1)))
        End Select
    End If
End Sub

Public Sub listPlayer(Data As String)
    lvwPlayer.ListItems.Clear
    loopX = 0
    Do
        If Split(Data, ";")(loopX) <> "EOT" Then
            lvwPlayer.ListItems.Add(, , Split(Split(Data, ";")(loopX), ",")(0)).SubItems(1) = Split(Split(Data, ";")(loopX), ",")(1)
        End If
        loopX = loopX + 1
    Loop Until Split(Data, ";")(loopX - 1) = "EOT"
    wskClient.SendData "req;tList"
    DoEvents
End Sub

Public Sub listTable(Data As String)
    lvwGameTable.ListItems.Clear
    loopX = 0
    Do
        If Split(Data, ";")(loopX) <> "EOT" Then
            With lvwGameTable.ListItems.Add(, , Split(Split(Data, ";")(loopX), ",")(0))
                .SubItems(1) = Split(Split(Data, ";")(loopX), ",")(1)
                If Split(Split(Data, ";")(loopX), ",")(2) = 0 Then
                    .SubItems(2) = "Original"
                ElseIf Split(Split(Data, ";")(loopX), ",")(2) = 1 Then
                    .SubItems(2) = "Short Game"
                Else
                    .SubItems(2) = "Custom"
                End If
                .SubItems(3) = Split(Split(Data, ";")(loopX), ",")(3)
                .SubItems(4) = Split(Split(Data, ";")(loopX), ",")(4)
                .SubItems(5) = Split(Split(Data, ";")(loopX), ",")(5)
                .SubItems(6) = Split(Split(Data, ";")(loopX), ",")(6)
                .SubItems(7) = Split(Split(Data, ";")(loopX), ",")(7)
                .SubItems(8) = Split(Split(Data, ";")(loopX), ",")(8)
            End With
        End If
        loopX = loopX + 1
    Loop Until Split(Data, ";")(loopX - 1) = "EOT"
End Sub

Public Sub addRemoveToList(Data As String)
    loopX = 0
    Do
        If Split(Data, ";")(loopX) <> "EOT" Then
            If Split(Split(Data, ";")(loopX), ",")(0) = "addToList" Then
                    lvwPlayer.ListItems.Add(, , Split(Split(Data, ";")(loopX), ",")(1)).SubItems(1) = Split(Split(Data, ";")(loopX), ",")(2)
            ElseIf Split(Split(Data, ";")(loopX), ",")(0) = "removeFromList" Then
                    On Error Resume Next
                    lvwPlayer.ListItems.Remove lvwPlayer.FindItem(Split(Split(Data, ";")(loopX), ",")(1)).Index
            ElseIf Split(Split(Data, ";")(loopX), ",")(0) = "create" Then
                With lvwGameTable.ListItems.Add(, , Split(Split(Data, ";")(loopX), ",")(1))
                    .SubItems(1) = Split(Split(Data, ";")(loopX), ",")(2)
                    If Split(Split(Data, ";")(loopX), ",")(3) = 0 Then
                        .SubItems(2) = "Original"
                    ElseIf Split(Split(Data, ";")(loopX), ",")(3) = 1 Then
                        .SubItems(2) = "Short Game"
                    Else
                        .SubItems(2) = "Custom"
                    End If
                    .SubItems(3) = Split(Split(Data, ";")(loopX), ",")(4)
                    .SubItems(4) = Split(Split(Data, ";")(loopX), ",")(5)
                    .SubItems(5) = Split(Split(Data, ";")(loopX), ",")(6)
                    .SubItems(6) = Split(Split(Data, ";")(loopX), ",")(7)
                    .SubItems(7) = Split(Split(Data, ";")(loopX), ",")(8)
                    .SubItems(8) = Split(Split(Data, ";")(loopX), ",")(9)
                End With
            ElseIf Split(Split(Data, ";")(loopX), ",")(0) = "remove" Then
                On Error Resume Next
                lvwGameTable.ListItems.Remove lvwGameTable.FindItem(Split(Split(Data, ";")(loopX), ",")(1)).Index
            End If
        End If
        loopX = loopX + 1
    Loop Until Split(Data, ";")(loopX - 1) = "EOT"
End Sub

