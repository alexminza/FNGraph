VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Public Expression As New clsExpression
Public Description As String
Public Min As Currency
Public Max As Currency
Public Precision As Currency
Public MaxGap As Long
Public Color As Long
Public DrawLines As Boolean
Public DrawPoints As Boolean
Public LinesWidth As Long
Public PointsSize As Long
Public Visible As Boolean
Public ValuesCount As Long

Private EvVals() As Currency
Private EvStats() As Long

Public Function DisplayName() As String
    If Description <> "" Then
        DisplayName = Expression.Expression & " [" & Description & "]"
    Else
        DisplayName = Expression.Expression
    End If
End Function

Public Sub CacheValues()
    On Error GoTo ErrHandler

    Dim L As Long
    L = (Max - Min + 1) * (1 / Precision) + 1
    ReDim EvVals(1 To 2, 1 To L)
    ReDim EvStats(1 To L)

    'Caching values
    Dim X As Currency, T As Long, EvStat As Long

    For X = Min To Max Step Precision
        T = T + 1
        EvVals(1, T) = X
        EvVals(2, T) = Expression.EvalFn(X, EvStat)
        EvStats(T) = EvStat
    Next X

    ValuesCount = T
    Exit Sub

ErrHandler:
    If ErrAssist(vbRetryCancel) = vbRetry Then Resume
End Sub

Public Sub EvVal(Index As Long, ByRef X As Currency, ByRef Y As Currency, ByRef EvStat As Long)
    X = EvVals(1, Index)
    Y = EvVals(2, Index)
    EvStat = EvStats(Index)
End Sub
