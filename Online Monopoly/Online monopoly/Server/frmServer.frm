VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Online Monopoly - Server"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
   Begin VB.Timer tmrSendMsg 
      Enabled         =   0   'False
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer tmrServerStatus 
      Interval        =   1000
      Left            =   960
      Top             =   0
   End
   Begin MSWinsockLib.Winsock wskListeningServer 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   5200
   End
   Begin VB.Label lblServerState 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   4695
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Set db = New Connection
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb;"
    Me.Caption = App.title
    wskListeningServer.LocalPort = 5155
    wskListeningServer.Listen
    Text1.Text = "Server has started." & vbCrLf
    Text1.Text = Text1.Text & "Server IP: " & wskListeningServer.LocalIP & vbCrLf
    Text1.Text = Text1.Text & "Server Port: " & wskListeningServer.LocalPort & vbCrLf
    Text1.SelStart = Len(frmServer.Text1.Text)
End Sub


Private Sub tmrServerStatus_Timer()
    lblServerState.Caption = "Server: " & GetStatus(wskListeningServer.state)
    For serverNo = 1 To wskServer.ubound
        If wskServer(serverNo).state <> sckConnected And player(serverNo).name <> "" Then
            Call addRemoveToList("removeFromList", player(serverNo).userID & "," & player(serverNo).name)
            If player(serverNo).inGame Then
                player(serverNo).inGame = False
            End If
            player(serverNo).userID = 0
            player(serverNo).name = ""
            wskServer(serverNo).Close
        End If
    Next
End Sub

Private Sub wskServer_DataArrival(index As Integer, ByVal bytesTotal As Long)
    Dim strMsg As String
    wskServer(index).GetData strMsg
    If strMsg <> "" Then
        receiveData index, strMsg
    End If
End Sub

Private Sub wskListeningServer_ConnectionRequest(ByVal requestID As Long)
    Dim freeServer As Integer
    If wskListeningServer.state = sckListening Then
        freeServer = 0
        For serverNo = 1 To wskServer.ubound
            If wskServer(serverNo).state <> sckConnected Then
                freeServer = serverNo
                Exit For
            End If
        Next
        If freeServer = 0 Then
            freeServer = wskServer.ubound + 1
            Load wskServer(freeServer)
        End If
        wskServer(freeServer).Close
        wskServer(freeServer).Accept requestID
        Do
            DoEvents
        Loop Until wskServer(freeServer).state = sckConnected
    End If
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

