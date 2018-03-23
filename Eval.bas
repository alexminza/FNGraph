Attribute VB_Name = "modEval"

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Public Const MATH_E As Currency = 2.71828182845905
Public Const MATH_PI As Currency = 3.14159265358979
Public Const MATH_ATN1 As Currency = 0.785398163397448 'PI/4
Public Const MATH_2ATN1 As Currency = 1.5707963267949 'PI/2
Public Const MATH_LOG10 As Currency = 2.30258509299405
'Public Const MATH_PI180 As Currency = 1.74532925199433E-02 'PI/180

Public Const ERR_UNKNOWNID = vbObjectError + 1

Public Function Min(A As Currency, B As Currency) As Currency
    If A < B Then
        Min = A
    Else
        Min = B
    End If
End Function

Public Function Max(A As Currency, B As Currency) As Currency
    If A > B Then
        Max = A
    Else
        Max = B
    End If
End Function

Public Function Fact(Number As Currency) As Currency
    If Number < 0 Then
        Err.Raise 13
    Else
        Dim I As Long, Result As Currency

        Result = 1
        For I = 2 To Fix(Number)
            Result = Result * I
        Next I

        Fact = Result
    End If
End Function

Public Function IsNumber(Buffer As String) As Boolean
    Dim I As Long, S As Long, DecSep As Boolean

    Select Case Buffer
        Case "", "-", "."
            IsNumber = False
            Exit Function
    End Select

    If Right(Buffer, 1) = "." Then
        IsNumber = False
        Exit Function
    End If

    If Left(Buffer, 1) = "-" Then
        S = 2
    Else
        S = 1
    End If

    For I = S To Len(Buffer)
        Select Case Mid(Buffer, I, 1)
            Case "0" To "9"
            Case "."
                If DecSep Then
                    IsNumber = False
                    Exit Function
                End If
                DecSep = True
            Case Else
                IsNumber = False
                Exit Function
        End Select
    Next I

    IsNumber = True
End Function

Public Function CurToStr(Number As Currency) As String
    Dim Buffer As String

    Buffer = Str(Number)
    If Left(Buffer, 1) = " " Then
        Buffer = Right(Buffer, Len(Buffer) - 1)
    End If
    If Left(Buffer, 1) = "." Then
        Buffer = "0" & Buffer
    End If
    If Left(Buffer, 2) = "-." Then
        Buffer = "-0." & Right(Buffer, Len(Buffer) - 2)
    End If

    CurToStr = Buffer
End Function

Public Function LngToStr(Number As Long) As String
    Dim Buffer As String

    Buffer = Str(Number)
    If Left(Buffer, 1) = " " Then
        LngToStr = Right(Buffer, Len(Buffer) - 1)
    Else
        LngToStr = Buffer
    End If
End Function

Public Function StrToCur(Buffer As String) As Currency
    If IsNumber(Buffer) Then
        StrToCur = Val(Buffer)
    Else
        Err.Raise 13
    End If
End Function

Public Function StrToLng(Buffer As String) As Long
    If IsNumber(Buffer) Then
        StrToLng = Val(Buffer)
    Else
        Err.Raise 13
    End If
End Function

Public Function LenExprSep(Expression As String) As Long
    Dim I As Long, Depth As Integer

    For I = 1 To Len(Expression)
        Select Case Mid(Expression, I, 1)
            Case "(": Depth = Depth + 1
            Case ")": Depth = Depth - 1
        End Select

        If Depth = 0 Then
            LenExprSep = I
            Exit Function
        End If

        Debug.Assert Depth >= 0 'The parentheses are always checked in the first phase. Just in any case...
    Next I

    Debug.Assert False 'It should never get here!
End Function

Public Function PrepareExpression(Expression As String)
    Expression = LCase(Expression)

    Dim I As Long, Ch As String, Result As String
    For I = 1 To Len(Expression)
        Ch = Mid(Expression, I, 1)

        Select Case Ch
            Case " "
            Case Else
                Result = Result & Ch
        End Select
    Next I

    PrepareExpression = Result
End Function

Public Function CheckParentheses(Expression As String) As Boolean
'returns True in case of error
    Dim I As Long, Depth As Long, LastPos As Long

    For I = 1 To Len(Expression)
        Select Case Mid(Expression, I, 1)
            Case "("
                LastPos = I
                Depth = Depth + 1
            Case ")"
                If LastPos = I - 1 Then 'excluding expressions with empty parenthesis ()
                    CheckParentheses = True
                    Exit Function
                End If
                Depth = Depth - 1
        End Select

        If Depth < 0 Then
            CheckParentheses = True
            Exit Function
        End If
    Next I

    CheckParentheses = (Depth <> 0)
End Function

Public Function CheckExpressionSyntax(Expression As String, Variables As clsVariablesCollection)
'returns True in case of error
    If CheckParentheses(Expression) Then
        MsgBox "Illegal Function value. Check parentheses.", vbExclamation
        CheckExpressionSyntax = True
        Exit Function
    End If

    Dim TestExpression As New clsExpression
    TestExpression.Expression = Expression
    Set TestExpression.Variables = Variables

    If TestExpression.Build Then
        'there was an error building the expression tree
        CheckExpressionSyntax = True
    End If
End Function

Public Function CheckVariableName(Name As String)
'returns True in case of error
    Select Case Left(Name, 1)
        Case "a" To "z"
        Case Else
            MsgBox "Invalid Name value. Must begin with a letter.", vbExclamation
            CheckVariableName = True
            Exit Function
    End Select

    Dim I As Long
    For I = 2 To Len(Name)
        Select Case Mid(Name, I, 1)
            Case "a" To "z", "0" To "9"
            Case Else
                MsgBox "Invalid Name value. Must consist only of letters or digits.", vbExclamation
                CheckVariableName = True
                Exit Function
        End Select
    Next I

    'check if it is a reserved id
    Select Case Name
        Case "x", "e", "pi", "min", "max", "mod", "radians", "degrees", _
        "sin", "cos", "tan", "atn", "abs", "ln", "lg", "log", "exp", "sqr", "fix", "int", "sgn", _
        "sec", "cosec", "cotan", "arcsin", "arccos", "arcsec", "arccosec", "arccotan", _
        "hsin", "hcos", "htan", "hsec", "hcosec", "hcotan", "harcsin", "harccos", "harctan", "harcsec", "harccosec", "harccotan", _
        "arctan", "arctg", "atan", "acot", "arccot", "asin", "acos", "acsc", "arccsc", "asec", "sinh", "sh", "cosh", "ch", "tanh", "th", _
        "coth", "cth", "sech", "csch", "acsch", "arccsch", "asech", "arcsech", "asinh", "arcsinh", "acosh", "arccosh", "atanh", "arctanh", _
        "acoth", "arccoth", "tg", "ctg", "csc", "sqrt", "trunc", "sign"
            MsgBox "Illegal Name value. Reserved identifier.", vbExclamation
            CheckVariableName = True
            Exit Function
    End Select
End Function
