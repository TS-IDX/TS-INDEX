VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsMemos"
Attribute VB_Ext_KEY = "Member0" ,"clsMemos"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(Target As String, Source As String, Message As String, Optional sKey As String) As clsMemos
Dim objNewMember As clsMemos
Set objNewMember = New clsMemos
objNewMember.MemoID = sKey
objNewMember.Message = Message
objNewMember.Source = Source
objNewMember.Target = Target
If Len(sKey) = 0 Then
    mCol.Add objNewMember
Else
    mCol.Add objNewMember, sKey
End If
Set Add = objNewMember
Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As clsMemos
Attribute Item.VB_UserMemId = 0
Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
mCol.Remove vntIndexKey
End Sub

Private Sub Class_Initialize()
Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
Set mCol = Nothing
End Sub

