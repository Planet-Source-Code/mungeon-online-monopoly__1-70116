Attribute VB_Name = "modSendData"
'##### login
'##### message
Public strPublicMsg As String
'##### game info
'##### cmd
Public strCmd As String

Public Sub addRemoveToList(cmd As String, data As String)
    Dim serverNo As Integer
    For serverNo = 1 To frmServer.wskServer.ubound
        If frmServer.wskServer(serverNo).state = sckConnected And Not player(serverNo).inGame Then
            frmServer.wskServer(serverNo).sendData "tpList;" & cmd & "," & data & ";" & "EOT;"
            DoEvents
        End If
    Next
End Sub
