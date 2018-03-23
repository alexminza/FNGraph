VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Public Name As String
Public Value As Currency

Public Function DisplayName() As String
    DisplayName = Name & " = " & CurToStr(Value)
End Function
