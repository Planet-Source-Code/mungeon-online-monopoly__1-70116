Attribute VB_Name = "mod1"
Public db As Connection
Public ado As Recordset
Public Const maxTable = 25
Public Const maxPlayer = 100
Public Type TableInfo
    TableId As Integer
    title As String
    password As String
    maxPlayer As Integer
    typeOfRules As Integer
    createdBy As Integer
    player(1 To 4) As Integer
    gameStarted As Boolean
    timeStart As Date
End Type

Public Type PlayerInfo
    userID As Long
    name As String
    strMsg As String
    inGame As Boolean
End Type

Public Table(0 To maxTable) As TableInfo
Public player(0 To maxPlayer) As PlayerInfo

Public strMsgAll As String
Public lineMsgAll As Integer

