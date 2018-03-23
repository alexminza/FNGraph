VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVariablesCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Public VariablesCollection As New Collection 'of clsVariable

Public Function IsDefined(Name As String) As Boolean
    Dim I As Long

    For I = 1 To VariablesCollection.Count
        If VariablesCollection(I).Name = Name Then
            IsDefined = True
            Exit Function
        End If
    Next I
End Function

Public Function GetValue(Name As String) As Currency
    Dim I As Long

    For I = 1 To VariablesCollection.Count
        If VariablesCollection(I).Name = Name Then
            GetValue = VariablesCollection(I).Value
            Exit Function
        End If
    Next I

    Debug.Assert False 'The variable should always be defined
End Function
