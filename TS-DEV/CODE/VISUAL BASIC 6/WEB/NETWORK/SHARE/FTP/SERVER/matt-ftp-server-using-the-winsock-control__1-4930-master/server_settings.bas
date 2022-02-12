Attribute VB_Name = "server_settings"
Option Explicit

''''''''''''''''''''''''''''
'SERVER SETTINGS
''''''''''''''''''''''''''''

Public Server_Name As String
Public Server_Port As Integer
Public Server_Welcome_Msg As String

'I doubt this server could even handle 10 clients. :)
Public Const Server_Max_Clients As Integer = 100
