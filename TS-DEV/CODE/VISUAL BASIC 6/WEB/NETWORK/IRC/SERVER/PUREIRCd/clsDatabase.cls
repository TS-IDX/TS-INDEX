VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private mvarFileName As String
Private mvarItemsPerEntry As Long

Public Property Let ItemsPerEntry(ByVal Items As Long)
Attribute ItemsPerEntry.VB_Description = "Let's you Specify how many Items will be listed under each Item"
    mvarItemsPerEntry = Items
End Property

Public Property Get ItemsPerEntry() As Long
    ItemsPerEntry = mvarItemsPerEntry
End Property

Public Property Let FileName(ByVal FileName As String)
Attribute FileName.VB_Description = "Specify a FileName to be used when reading/writing"
    mvarFileName = FileName
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Sub WriteEntry(Section As String, Key As String, Items As String)
WritePrivateProfileString Section, Key, Items, mvarFileName
End Sub

Public Function ReadEntry(Section As String, Key As String, Default As String) As String
Dim buf As String
buf = Space$(255)
GetPrivateProfileString Section, Key, Default, buf, Len(buf), mvarFileName
ReadEntry = StripTerminator(buf)
End Function

Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
