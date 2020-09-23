Attribute VB_Name = "mod2"
Public Function setPlayerStatus(status As String, Index As Integer, value As Variant)
    Select Case LCase(status)
        Case "pid"
            player(Index).PID = value
        Case "pname"
            player(Index).PName = value
            Monopoly.lblPlayerName(Index).Caption = player(Index).PName
        Case "cash"
            If value <> player(Index).cash Then
                Monopoly.lblCashFlow(Index).Caption = "$" & value - player(Index).cash
            Else
                Monopoly.lblCashFlow(Index).Caption = ""
            End If
            player(Index).cash = value
            Monopoly.lblPlayerCash(Index).Caption = "$" & player(Index).cash
            If Index = PlayerNumber Then
                Monopoly.lblCurCash.Caption = "$" & player(Index).cash
            End If
        Case "tokenid"
            player(Index).tokenID = value
            Monopoly.imgPlayerToken(Index).Picture = LoadPicture("images/Token/" & token(value).file)

    End Select
End Function
