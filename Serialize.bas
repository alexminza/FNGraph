Attribute VB_Name = "modSerialize"

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Global Const FNG_TAG_GRAPH = "graph", FNG_TAG_VARIABLE = "variable", FNG_TAG_AXES = "axes", FNG_TAG_GRID = "grid", FNG_TAG_VALUES = "values"

Public Function XMLFindWhiteSpace(Buffer As String, Optional Start = 1) As Long
    Dim I As Long
    
    For I = Start To Len(Buffer)
        Select Case Mid(Buffer, I, 1)
            Case " ", vbTab, vbCr, vbLf
                XMLFindWhiteSpace = I
                Exit Function
        End Select
    Next I

    XMLFindWhiteSpace = 0
End Function

Public Function StrToBln(Buffer As String) As Boolean
    Select Case LCase(Buffer)
        Case "true"
            StrToBln = True
        Case "false"
            StrToBln = False
    End Select
End Function

Public Function BlnToStr(Value As Boolean) As String
    If Value Then
        BlnToStr = "True"
    Else
        BlnToStr = "False"
    End If
End Function

Public Function ColorToStr(Value As Long) As String
    Dim HexVal As String, I As Long

    HexVal = Hex(Value)
    For I = Len(HexVal) + 1 To 6
        HexVal = "0" & HexVal
    Next I

    ColorToStr = "#" & Right(HexVal, 2) & Mid(HexVal, 3, 2) & Left(HexVal, 2)
End Function

Public Function StrToColor(Buffer As String) As Long
    Dim HexVal As String

    HexVal = Right(Buffer, 6) 'cutting the leading "#"
    HexVal = "&H" & Right(HexVal, 2) & Mid(HexVal, 3, 2) & Left(HexVal, 2)
    StrToColor = CLng(HexVal)
End Function
