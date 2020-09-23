Attribute VB_Name = "modReceiveData"
Public serverNo As Integer
Dim loopX As Integer

Public Function receiveData(index As Integer, strMsg As String)
    Select Case Split(strMsg, ";")(0)
        Case "register"
            Set ado = New Recordset
            ado.Open "SELECT * FROM tblUser WHERE Username='" & Split(strMsg, ";")(1) & "'", db, adOpenStatic, adLockOptimistic
            If ado.EOF Then
                ado.AddNew
                ado!UserName = Split(strMsg, ";")(1)
                ado!password = Split(strMsg, ";")(2)
                ado.Update
                frmServer.wskServer(index).SendData "register;successful;"
                DoEvents
            Else
                frmServer.wskServer(index).SendData "register;failed;"
                DoEvents
            End If
            Set ado = Nothing
        Case "login"
            Dim strLoginStatus As String
            strLoginStatus = ""
            For loopX = 0 To maxPlayer
                If Split(strMsg, ";")(1) = player(loopX).name Then
                    strLoginStatus = "login;InvalidLogin"
                End If
            Next
            If strLoginStatus <> "login;InvalidLogin" Then
                Set ado = New Recordset
                ado.Open "SELECT * FROM tblUser WHERE Username='" & Split(strMsg, ";")(1) & "'", db, adOpenStatic, adLockOptimistic
                If Not ado.EOF Then
                    If ado!password = Split(strMsg, ";")(2) Then
                        player(index).userID = ado!userID
                        player(index).name = ado!UserName
                        player(index).strMsg = ""
                        player(index).inGame = False
                        strLoginStatus = "login;Successful;" & ado!userID & "," & ado!UserName
                        Call addRemoveToList("addToList", player(index).userID & "," & player(index).name)
                        frmServer.Text1.Text = frmServer.Text1.Text & player(index).name & " has logon. " & vbCrLf
                        frmServer.Text1.SelStart = Len(frmServer.Text1.Text)
                    Else
                        strLoginStatus = "login;InvalidPassword"
                    End If
                Else
                    strLoginStatus = "login;NotExist"
                End If
                Set ado = Nothing
            End If
            frmServer.wskServer(index).SendData strLoginStatus & ";"
            DoEvents
        Case "logout"
            Call addRemoveToList("removeFromList", player(index).userID & "," & player(index).name)
            frmServer.Text1.Text = frmServer.Text1.Text & player(index).name & " has logout. " & vbCrLf
            frmServer.Text1.SelStart = Len(frmServer.Text1.Text)
            player(index).userID = 0
            player(index).name = ""
            frmServer.wskServer(index).Close
        Case "msg"
                If Split(strMsg, ";")(1) = "public" Then
                    For serverNo = 1 To frmServer.wskServer.ubound
                        If frmServer.wskServer(serverNo).state = sckConnected And Not player(serverNo).inGame Then
                            strPublicMsg = player(index).name & ": " & Right$(strMsg, Len(strMsg) - (Len(Split(strMsg, ";")(0)) + Len(Split(strMsg, ";")(1)) + 2))
                            frmServer.wskServer(serverNo).SendData "msg;" & strPublicMsg
                            DoEvents
                        End If
                    Next
                ElseIf Split(strMsg, ";")(1) = "table" Then
                    For serverNo = 1 To Table(Int(Split(strMsg, ";")(2))).maxPlayer
                        If Table(Int(Split(strMsg, ";")(2))).player(serverNo) <> 0 Then
                            strPublicMsg = player(index).name & ": " & Right$(strMsg, Len(strMsg) - (Len(Split(strMsg, ";")(0)) + Len(Split(strMsg, ";")(1)) + Len(Split(strMsg, ";")(2)) + Len(Split(strMsg, ";")(3)) + 4))
                            frmServer.wskServer(Table(Int(Split(strMsg, ";")(2))).player(serverNo)).SendData "tableMsg;" & strPublicMsg
                            DoEvents
                        End If
                    Next
                ElseIf Split(strMsg, ";")(1) = "player" Then
'                    If player(serverNo).strMsg <> "" Then
'                        wskServer(serverNo).sendData "pMsg;" & player(serverNo).strMsg & "EOT;"
'                        DoEvents
'                    End If
                End If
        Case "table"
            If Split(strMsg, ";")(1) = "create" Then
                Dim id As Integer
                For id = 1 To maxTable
                    If Table(id).TableId = 0 Then
                        Exit For
                    End If
                Next
                If Table(id).TableId = 0 Then
                    Table(id).TableId = id
                    Table(id).title = Split(strMsg, ";")(2)
                    Table(id).maxPlayer = Split(strMsg, ";")(3)
                    Table(id).typeOfRules = Split(strMsg, ";")(4)
                    Table(id).gameStarted = False
                    Table(id).createdBy = index
                    Table(id).player(1) = index
                    Call addRemoveToList("create", id & "," & Table(id).title & "," & Table(id).typeOfRules & "," & Table(id).maxPlayer & "," & player(Table(id).player(1)).name & "," & player(Table(id).player(2)).name & "," & player(Table(id).player(3)).name & "," & player(Table(id).player(4)).name & ",Waiting")
                    frmServer.wskServer(index).SendData "setTableID;" & id & ";"
                    DoEvents
                Else
                    MsgBox "Table are Fulled"
                End If
            ElseIf Split(strMsg, ";")(1) = "reqGameInfo" Then
                For loopX = 1 To 4
                    If Table(Split(strMsg, ";")(2)).player(loopX) <> 0 Then
                        frmServer.wskServer(Table(Split(strMsg, ";")(2)).player(loopX)).SendData "reqGameInfo;" & index & ";" & player(index).userID & ";" & player(index).name & ";"
                        DoEvents
                        Exit For
                    End If
                Next
            ElseIf Split(strMsg, ";")(1) = "gameInfo" Then
                frmServer.wskServer(Int(Split(strMsg, ";")(2))).SendData "receiveGameInfo;" & Split(strMsg, ";")(3) & ";"
                DoEvents
            ElseIf Split(strMsg, ";")(1) = "startGame" Then
                For loopX = 1 To 4
                    If Table(Int(Split(strMsg, ";")(2))).player(loopX) <> 0 Then
                        frmServer.wskServer(Table(Split(strMsg, ";")(2)).player(loopX)).SendData "startGame;" & Right$(strMsg, Len(strMsg) - (Len(Split(strMsg, ";")(0)) + Len(Split(strMsg, ";")(1)) + Len(Split(strMsg, ";")(2)) + 3))
                        DoEvents
                    End If
                Next
                Table(Int(Split(strMsg, ";")(2))).gameStarted = True
                For serverNo = 1 To frmServer.wskServer.ubound
                    If frmServer.wskServer(serverNo).state = sckConnected And Not player(serverNo).inGame Then
                        frmServer.wskServer(serverNo).SendData "updateStartedGame;" & Split(strMsg, ";")(1) & ";"
                        DoEvents
                    End If
                Next
            ElseIf Split(strMsg, ";")(1) = "join" Then
                Dim tempLoc As Integer
                Dim strJoinMsg As String
                For loopX = 1 To Table(Split(strMsg, ";")(2)).maxPlayer
                    If Table(Split(strMsg, ";")(2)).player(loopX) = 0 Then
                        tempLoc = loopX
                        Table(Split(strMsg, ";")(2)).player(loopX) = index
                        Exit For
                    End If
                Next
                For loopX = 1 To Table(Split(strMsg, ";")(2)).maxPlayer
                    If Table(Split(strMsg, ";")(2)).player(loopX) <> 0 Then
                        For serverNo = 1 To frmServer.wskServer.ubound
                            If frmServer.wskServer(serverNo).state = sckConnected And Not player(serverNo).inGame Then
                                If Table(Split(strMsg, ";")(2)).player(loopX) = serverNo Then
                                    strJoinMsg = "playerJoin;" & tempLoc & ";" & player(index).userID & ";" & player(index).name
                                    frmServer.wskServer(serverNo).SendData strJoinMsg & ";EOT;"
                                    DoEvents
                                Else
                                    frmServer.wskServer(serverNo).SendData "updateTList;" & Table(Split(strMsg, ";")(2)).TableId & ";" & tempLoc & ";" & player(index).userID & ";" & player(index).name & ";EOT;"
                                    DoEvents
                                End If
                            ElseIf frmServer.wskServer(serverNo).state = sckConnected And player(serverNo).inGame Then
                                If serverNo = Table(Split(strMsg, ";")(2)).player(1) Or serverNo = Table(Split(strMsg, ";")(2)).player(2) Or serverNo = Table(Split(strMsg, ";")(2)).player(3) Or serverNo = Table(Split(strMsg, ";")(2)).player(4) Then
                                    strJoinMsg = "playerJoin;" & tempLoc & ";" & player(index).userID & ";" & player(index).name
                                    frmServer.wskServer(serverNo).SendData strJoinMsg & ";EOT;"
                                    DoEvents
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        Case "status"
            Dim tableDeleted As Boolean
            tableDeleted = False
            If Split(strMsg, ";")(1) = "enterGame" Then
                player(index).inGame = True
                Call addRemoveToList("removeFromList", player(index).userID & "," & player(index).name)
            ElseIf Split(strMsg, ";")(1) = "quitGame" Then
                Dim strTemp As String
                player(index).inGame = False
                strTemp = player(index).userID & "," & player(index).name
                For loopX = 1 To 4
                    If Int(Split(strMsg, ";")(3)) = loopX Then Table(Split(strMsg, ";")(2)).player(loopX) = 0
                Next
                If Table(Split(strMsg, ";")(2)).player(1) = 0 And Table(Split(strMsg, ";")(2)).player(2) = 0 And Table(Split(strMsg, ";")(2)).player(3) = 0 And Table(Split(strMsg, ";")(2)).player(4) = 0 Then
                    Table(Split(strMsg, ";")(2)).TableId = 0
                    strTemp = strTemp & ";" & "remove" & "," & Split(strMsg, ";")(2) & "," & Table(Split(strMsg, ";")(2)).title
                    tableDeleted = True
                End If
                If strTemp <> "" Then
                    Call addRemoveToList("addToList", strTemp)
                End If
                For serverNo = 1 To frmServer.wskServer.ubound
                    If frmServer.wskServer(serverNo).state = sckConnected And Not player(serverNo).inGame Then
                        If index = serverNo Then
                            genPlist index
                        Else
                            If Not tableDeleted Then
                                frmServer.wskServer(serverNo).SendData "updateTList;" & Split(strMsg, ";")(2) & ";" & Split(strMsg, ";")(3) & ";;;EOT;"
                                DoEvents
                            End If
                        End If
                    ElseIf frmServer.wskServer(serverNo).state = sckConnected And player(serverNo).inGame Then
                        If serverNo = Table(Split(strMsg, ";")(2)).player(1) Or serverNo = Table(Split(strMsg, ";")(2)).player(2) Or serverNo = Table(Split(strMsg, ";")(2)).player(3) Or serverNo = Table(Split(strMsg, ";")(2)).player(4) Then
                            frmServer.wskServer(serverNo).SendData "playerQuit;" & Split(strMsg, ";")(3) & ";" & player(index).userID & ";" & player(index).name & ";EOT;"
                            DoEvents
                        End If
                    End If
                Next
            End If
        Case "req"
            If Split(strMsg, ";")(1) = "pList" Then
                genPlist index
            ElseIf Split(strMsg, ";")(1) = "tList" Then
                genTlist index
            End If
        Case "cmd"
            On Error Resume Next
            Dim strCmd As String
            For loopX = 1 To 4
                If Table(Int(Split(strMsg, ";")(2))).player(loopX) <> 0 Then
                    strCmd = Split(strMsg, ";")(1) & ";" & Right$(strMsg, Len(strMsg) - (Len(Split(strMsg, ";")(0)) + Len(Split(strMsg, ";")(1)) + Len(Split(strMsg, ";")(2)) + 3))
                    frmServer.wskServer(Table(Split(strMsg, ";")(2)).player(loopX)).SendData strCmd
                    DoEvents
                End If
            Next
    End Select
End Function

Public Sub genPlist(index As Integer)
    Dim strPlayerList As String
    For loopX = 1 To frmServer.wskServer.ubound
        If frmServer.wskServer(loopX).state = sckConnected And player(loopX).userID <> 0 And Not player(loopX).inGame Then
            strPlayerList = strPlayerList & player(loopX).userID & "," & player(loopX).name & ";"
        End If
    Next
    If strPlayerList <> "" Then
        frmServer.wskServer(index).SendData "pList;" & strPlayerList & "EOT;"
        DoEvents
        strPlayerList = ""
    End If
End Sub

Public Sub genTlist(index As Integer)
    Dim strTableList As String
    For loopX = 1 To maxTable
        If Table(loopX).TableId <> 0 Then
            strTableList = strTableList & loopX & "," & Table(loopX).title & "," & Table(loopX).typeOfRules & "," & Table(loopX).maxPlayer & "," & player(Table(loopX).player(1)).name & "," & player(Table(loopX).player(2)).name & "," & player(Table(loopX).player(3)).name & "," & player(Table(loopX).player(4)).name & ","
            If Table(loopX).gameStarted Then
                strTableList = strTableList & "Playing" & ";"
            Else
                strTableList = strTableList & "Waiting" & ";"
            End If
        End If
    Next
    If strTableList <> "" Then
        frmServer.wskServer(index).SendData "tList;" & strTableList & "EOT;"
        DoEvents
        strTableList = ""
    End If
End Sub
