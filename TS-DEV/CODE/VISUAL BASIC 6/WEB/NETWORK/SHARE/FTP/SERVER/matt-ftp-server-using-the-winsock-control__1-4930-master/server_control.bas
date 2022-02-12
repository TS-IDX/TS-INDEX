Attribute VB_Name = "server_control"
Option Explicit

Public connected_clients As Integer

''''''''''''''''''''''''''''
'START UP THE SERVER
''''''''''''''''''''''''''''
Public Function start_server()

    On Error GoTo HandleError

    Dim server_status As String

    'if user defined settings exist, load them here
    'else load defaults
    loadDefaults

    'This is the winsock object that all clients will go through first.
    frmMain.commandSock(0).LocalPort = Server_Port
    frmMain.commandSock(0).Listen

    server_status = "Server started - Listening on port " & Server_Port
    update_server_log server_status

HandleError:
    If Err.Number > 0 Then  'If there is an error, report it.

        If Err.Number = 10048 Then
            server_status = "ERROR - start_server: port " & Server_Port & " is in use."
        Else
            server_status = "ERROR - start_server: " & Err.Description & " (" & Err.Number & ")"
        End If
    
        update_server_log server_status

    End If

End Function

''''''''''''''''''''''''''''
'DEFAULT SETTINGS FOR SERVER
''''''''''''''''''''''''''''

Public Sub loadDefaults()

    Server_Name = "VB FTP Server ver .1a"
    Server_Port = 22
    Server_Welcome_Msg = Server_Name & " ready..."  'First message displayed when client connects.

End Sub

''''''''''''''''''''''''''''
'NEW CONNECTION
''''''''''''''''''''''''''''

Public Sub new_connection(ByVal requestID As Long)

    Dim socket As Integer
    Dim objTmp As Variant

    'First check for any unused sockets before creating a new one.
    For Each objTmp In frmMain.commandSock
        If frmMain.commandSock(objTmp.Index).State = sckClosed Then
            If objTmp.Index <> 0 Then socket = objTmp.Index
        End If
    Next

    If socket = 0 Then
        socket = frmMain.commandSock.Count
        'The reason for using .Count is because it the first instance (0), counts as 1.
        'So if the only instance you have is the first one, (0), then .Count returns 1.
        'Which means you can create a new winsock instance using 1.
        Load frmMain.commandSock(frmMain.commandSock.Count) 'new winsock instance for client to connect to.
        Load frmMain.dataSock(frmMain.dataSock.Count)    'allocate data socket for new client.
    End If

    frmMain.commandSock(socket).Accept requestID

    update_server_log "Connected to " & frmMain.commandSock(socket).RemoteHostIP & ", logging in..."

    send_response socket, "220 " & Server_Name  'send welcome message

End Sub

''''''''''''''''''''''''''''
'UPDATE SERVER LOG WINDOW
''''''''''''''''''''''''''''

Public Sub update_server_log(message As String)

    'Displays the specified message in the server log.
    frmMain.lstServerLog.AddItem Time & " - " & message

End Sub
