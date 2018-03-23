VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Public Expression As String
Public Variables As clsVariablesCollection

Private EvTree() As clsEvTreeNode, EvTreeNodesCount As Long

Public Function Build() As Boolean
    Debug.Assert Expression <> "" 'Was an expression set?

    Erase EvTree() 'TODO: Dereference all objects in array
    EvTreeNodesCount = 0
    AddEvTreeNode

    BuildBET 1, Expression
    PrepareBET 1
    Build = CheckBET(1) 'TODO: In case of fail - clean-up

    'NOTE: VB cleans the memory automatically (using object reference counter).
    'But it takes a bit before it does the garbage cleaning, so it's best to do it yourself.
End Function

Public Function Evaluate(X As Currency) As Currency
    Debug.Assert EvTreeNodesCount > 0 'Was an expression tree built?

    Evaluate = EvalBET(1, X)
End Function

Public Function EvalFn(X As Currency, EvStat As Long) As Currency
    On Error GoTo ErrHandler

    EvalFn = EvalBET(1, X)
    EvStat = 0
    Exit Function

ErrHandler:
    EvStat = EVAL_ERROR
End Function

Private Sub AddEvTreeNode()
    EvTreeNodesCount = EvTreeNodesCount + 1
    ReDim Preserve EvTree(1 To EvTreeNodesCount)

    Set EvTree(EvTreeNodesCount) = New clsEvTreeNode
End Sub

Private Sub BuildBET(EvTreeRoot As Long, ExpressionBuffer As String)
    Dim I As Long, P As Long, L As Long, Depth As Long, Op1 As String, Op2 As String

    'removing external enclosing braces
    L = Len(ExpressionBuffer)
    Do While Left(ExpressionBuffer, 1) = "("
        If LenExprSep(ExpressionBuffer) = L Then
            ExpressionBuffer = Mid(ExpressionBuffer, 2, L - 2)
            L = L - 2
        Else
            Exit Do
        End If
    Loop

    'is it a symbolic constant?
    Select Case ExpressionBuffer
        Case "x", "e", "pi": GoTo ProcExit
    End Select

    'is it a numeric constant?
    If IsNumber(ExpressionBuffer) Then GoTo ProcExit

    'separating operands
    Op1 = "+": Op2 = "-": GoSub Separation
    Op1 = "*": Op2 = "/": GoSub Separation
    Op1 = "^": Op2 = "": GoSub Separation
    Op1 = "!": Op2 = "": GoSub Separation

    'is it a function?
    I = InStr(ExpressionBuffer, "(")
    If I > 0 And Right(ExpressionBuffer, 1) = ")" Then 'it is a function
        EvTree(EvTreeRoot).Value = Left(ExpressionBuffer, I - 1)

        Depth = 0
        For P = I + 1 To L - 1
            Select Case Mid(ExpressionBuffer, P, 1)
                Case "(": Depth = Depth + 1
                Case ")": Depth = Depth - 1
                Case "," ' arguments separator
                    If Depth = 0 Then
                        If P > I + 1 Then
                            AddEvTreeNode
                            EvTree(EvTreeRoot).Left = EvTreeNodesCount
                            BuildBET EvTree(EvTreeRoot).Left, Mid(ExpressionBuffer, I + 1, P - I - 1)
                        Else
                            'if this tree is not new, there may be some old data
                            EvTree(EvTreeRoot).Left = 0
                        End If
                        If L > P + 1 Then
                            AddEvTreeNode
                            EvTree(EvTreeRoot).Right = EvTreeNodesCount
                            BuildBET EvTree(EvTreeRoot).Right, Mid(ExpressionBuffer, P + 1, L - P - 1)
                        Else
                            EvTree(EvTreeRoot).Right = 0
                        End If
                        Exit Sub
                    End If
            End Select
        Next P

        'single-argument function
        EvTree(EvTreeRoot).Left = 0

        If L > I + 1 Then
            AddEvTreeNode
            EvTree(EvTreeRoot).Right = EvTreeNodesCount
            BuildBET EvTree(EvTreeRoot).Right, Mid(ExpressionBuffer, I + 1, L - I - 1)
        Else
            EvTree(EvTreeRoot).Right = 0
        End If
        Exit Sub
    End If

ProcExit:
    EvTree(EvTreeRoot).Value = ExpressionBuffer
    EvTree(EvTreeRoot).Left = 0
    EvTree(EvTreeRoot).Right = 0
    Exit Sub

'find and separate operands
Separation:
    Depth = 0
    For I = L To 1 Step -1
        Select Case Mid(ExpressionBuffer, I, 1)
            Case "(": Depth = Depth + 1
            Case ")": Depth = Depth - 1
            Case Op1, Op2
                If Depth = 0 Then
                    EvTree(EvTreeRoot).Value = Mid(ExpressionBuffer, I, 1) 'one character operators

                    If I > 1 Then 'do we have the left operand?
                        AddEvTreeNode
                        EvTree(EvTreeRoot).Left = EvTreeNodesCount
                        BuildBET EvTree(EvTreeRoot).Left, Left(ExpressionBuffer, I - 1)
                    Else
                        EvTree(EvTreeRoot).Left = 0
                    End If
                    If I < L Then 'do we have the right operand?
                        AddEvTreeNode
                        EvTree(EvTreeRoot).Right = EvTreeNodesCount
                        BuildBET EvTree(EvTreeRoot).Right, Right(ExpressionBuffer, L - I)
                    Else
                        EvTree(EvTreeRoot).Right = 0
                    End If
                    Exit Sub
                End If
        End Select
    Next I
    Return
End Sub

Private Sub PrepareBET(EvTreeRoot As Long)
    If EvTree(EvTreeRoot).Left > 0 Then PrepareBET EvTree(EvTreeRoot).Left
    If EvTree(EvTreeRoot).Right > 0 Then PrepareBET EvTree(EvTreeRoot).Right

    Select Case EvTree(EvTreeRoot).Value
        Case "arctan", "arctg", "atan": EvTree(EvTreeRoot).Value = "atn"
        Case "acot", "arccot": EvTree(EvTreeRoot).Value = "arccotan"
        Case "asin": EvTree(EvTreeRoot).Value = "arcsin"
        Case "acos": EvTree(EvTreeRoot).Value = "arccos"
        Case "acsc", "arccsc": EvTree(EvTreeRoot).Value = "arccosec"
        Case "asec": EvTree(EvTreeRoot).Value = "arcsec"
        Case "sinh", "sh": EvTree(EvTreeRoot).Value = "hsin"
        Case "cosh", "ch": EvTree(EvTreeRoot).Value = "hcos"
        Case "tanh", "th": EvTree(EvTreeRoot).Value = "htan"
        Case "coth", "cth": EvTree(EvTreeRoot).Value = "hcotan"
        Case "sech": EvTree(EvTreeRoot).Value = "hsec"
        Case "csch": EvTree(EvTreeRoot).Value = "hcosec"
        Case "acsch", "arccsch": EvTree(EvTreeRoot).Value = "harccosec"
        Case "asech", "arcsech": EvTree(EvTreeRoot).Value = "harcsec"
        Case "asinh", "arcsinh": EvTree(EvTreeRoot).Value = "harcsin"
        Case "acosh", "arccosh": EvTree(EvTreeRoot).Value = "harccos"
        Case "atanh", "arctanh": EvTree(EvTreeRoot).Value = "harctan"
        Case "acoth", "arccoth": EvTree(EvTreeRoot).Value = "harccotan"
        Case "tg": EvTree(EvTreeRoot).Value = "tan"
        Case "ctg", "cot": EvTree(EvTreeRoot).Value = "cotan"
        Case "csc": EvTree(EvTreeRoot).Value = "cosec"
        Case "sqrt": EvTree(EvTreeRoot).Value = "sqr"
        Case "trunc": EvTree(EvTreeRoot).Value = "fix"
        Case "sign": EvTree(EvTreeRoot).Value = "sgn"
    End Select
End Sub

Private Function CheckBET(EvTreeRoot As Long) As Boolean
'Returns True when fails

    Dim Value As String, RLeft As Long, RRight As Long
    Value = EvTree(EvTreeRoot).Value
    RLeft = EvTree(EvTreeRoot).Left
    RRight = EvTree(EvTreeRoot).Right

    'numeric constant
    If IsNumber(Value) Then
        If (RLeft <> 0) Or (RRight <> 0) Then
            MsgBox "Not a function: " & Value, vbExclamation
            CheckBET = True
            Exit Function
        End If
        Exit Function
    End If

    Select Case Value
        'variables, constants
        Case "x", "e", "pi"
            If (RLeft <> 0) Or (RRight <> 0) Then
                MsgBox "Not a function: " & Value, vbExclamation
                CheckBET = True
                Exit Function
            End If

        'operators
        Case "+", "*", "/", "^"
            If (RLeft = 0) Or (RRight = 0) Then
                MsgBox "Expression expected: " & Value, vbExclamation
                CheckBET = True
                Exit Function
            End If

        Case "-"
            If RRight = 0 Then
                MsgBox "Expression expected: " & Value, vbExclamation
                CheckBET = True
                Exit Function
            End If

        Case "!"
            If RLeft = 0 Then
                MsgBox "Expression expected: " & Value, vbExclamation
                CheckBET = True
                Exit Function
            End If
            If RRight <> 0 Then
                MsgBox "Wrong number of operands: " & Value, vbExclamation
                CheckBET = True
                Exit Function
            End If

        'functions
        Case "abs", "arccos", "arccosec", "arccotan", "arcsec", "arcsin", "atn", "cos", "cosec", "cotan", _
        "exp", "fix", "harccos", "harccosec", "harccotan", "harcsec", "harcsin", "harctan", "hcos", "hcosec", _
        "hcotan", "hsec", "hsin", "htan", "int", "ln", "lg", "sec", "sgn", "sin", "sqr", "tan", "radians", "degrees"
            If RLeft <> 0 Then
                MsgBox "Wrong number of arguments: " & Value, vbExclamation
                CheckBET = True
                Exit Function
            End If
            If RRight = 0 Then
                MsgBox "Function argument missing: " & Value, vbExclamation
                CheckBET = True
                Exit Function
            End If

        'functions (2 arguments)
        Case "min", "max", "mod", "log"
            If (RLeft = 0) Or (RRight = 0) Then
                MsgBox "Function argument missing: " & Value, vbExclamation
                CheckBET = True
                Exit Function
            End If

        Case Else
            'isn't this a user-defined variable?
            If Not Variables.IsDefined(Value) Then
                MsgBox "Unknown identifier: " & Value, vbExclamation
                CheckBET = True
                Exit Function
            End If
    End Select

    'recurse "deeper"
    If RLeft > 0 Then
        If CheckBET(RLeft) Then
            CheckBET = True
            Exit Function
        End If
    End If
    If RRight > 0 Then
        If CheckBET(RRight) Then
            CheckBET = True
            Exit Function
        End If
    End If
End Function

Private Function EvalBET(EvTreeRoot As Long, X As Currency) As Currency
    Dim Value As String
    Value = EvTree(EvTreeRoot).Value

    'is it an intrinsic symbolic constant?
    Select Case Value
        Case "x": EvalBET = X: Exit Function
        Case "e": EvalBET = MATH_E: Exit Function
        Case "pi": EvalBET = MATH_PI: Exit Function
    End Select

    'is it a numeric constant?
    If IsNumber(Value) Then
        EvalBET = Val(Value)
        Exit Function
    End If

    '"dig deeper"
    Dim REval As Currency, LEval As Currency
    If EvTree(EvTreeRoot).Left > 0 Then LEval = EvalBET(EvTree(EvTreeRoot).Left, X)
    If EvTree(EvTreeRoot).Right > 0 Then REval = EvalBET(EvTree(EvTreeRoot).Right, X)

    Select Case Value
        Case "+": EvalBET = LEval + REval
        Case "-"
            If EvTree(EvTreeRoot).Left = 0 Then
                'no left operand, it is an unary minus
                EvalBET = -REval
            Else
                EvalBET = LEval - REval
            End If

        Case "*": EvalBET = LEval * REval
        Case "/": EvalBET = LEval / REval
        Case "^": EvalBET = LEval ^ REval
        Case "!": EvalBET = Fact(LEval)

        Case "min": EvalBET = Min(LEval, REval)
        Case "max": EvalBET = Max(LEval, REval)
        Case "mod": EvalBET = LEval Mod REval

        Case "radians": EvalBET = REval * MATH_PI / 180
        Case "degrees": EvalBET = REval * 180 / MATH_PI

        Case "sin": EvalBET = Sin(REval)
        Case "cos": EvalBET = Cos(REval)
        Case "tan": EvalBET = Tan(REval)
        Case "atn": EvalBET = Atn(REval)

        Case "abs": EvalBET = Abs(REval)
        Case "ln": EvalBET = Log(REval)
        Case "lg": EvalBET = Log(REval) / MATH_LOG10
        Case "log": EvalBET = Log(LEval) / Log(REval) 'Log(number,base)
        Case "exp": EvalBET = Exp(REval)
        Case "sqr": EvalBET = Sqr(REval)
        Case "fix": EvalBET = Fix(REval)
        Case "int": EvalBET = Int(REval)
        Case "sgn": EvalBET = Sgn(REval)

        'Sec(X) = 1 / Cos(X)
        Case "sec": EvalBET = 1 / Cos(REval)

        'Cosec(X) = 1 / Sin(X)
        Case "cosec": EvalBET = 1 / Sin(REval)

        'Cotan(X) = 1 / Tan(X)
        Case "cotan": EvalBET = 1 / Tan(REval)

        'Arcsin(X) = Atn(X / Sqr(-X * X + 1))
        Case "arcsin":
            Select Case REval
                Case -1: EvalBET = -MATH_2ATN1
                Case 1: EvalBET = MATH_2ATN1
                Case Else: EvalBET = Atn(REval / Sqr(-REval * REval + 1))
            End Select

        'Arccos(X) = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
        Case "arccos":
            Select Case REval
                Case -1: EvalBET = MATH_PI
                Case 1: EvalBET = 0
                Case Else: EvalBET = Atn(-REval / Sqr(-REval * REval + 1)) + MATH_2ATN1
            End Select

        'Arcsec(X) = Atn(X / Sqr(X * X – 1)) + Sgn((X) – 1) * (2 * Atn(1))
        Case "arcsec": EvalBET = Atn(REval / Sqr(REval * REval - 1)) + Sgn(REval - 1) * MATH_2ATN1

        'Arccosec(X) = Atn(X / Sqr(X * X - 1)) + (Sgn(X) – 1) * (2 * Atn(1))
        Case "arccosec": EvalBET = Atn(REval / Sqr(REval * REval - 1)) + (Sgn(X) - 1) * MATH_2ATN1

        'Arccotan(X) = Atn(X) + 2 * Atn(1)
        '!!! arccotan(x) = 2 * arctan(1) - arctan(x)
        Case "arccotan": EvalBET = MATH_2ATN1 - Atn(REval)

        'HSin(X) = (Exp(X) – Exp(-X)) / 2
        Case "hsin": EvalBET = (Exp(REval) - Exp(-REval)) / 2

        'HCos(X) = (Exp(X) + Exp(-X)) / 2
        Case "hcos": EvalBET = (Exp(REval) + Exp(-REval)) / 2

        'HTan(X) = (Exp(X) – Exp(-X)) / (Exp(X) + Exp(-X))
        Case "htan": EvalBET = (Exp(REval) - Exp(-REval)) / (Exp(REval) + Exp(-REval))

        'HSec(X) = 2 / (Exp(X) + Exp(-X))
        Case "hsec": EvalBET = 2 / (Exp(REval) + Exp(-REval))

        'HCosec(X) = 2 / (Exp(X) – Exp(-X))
        Case "hcosec": EvalBET = 2 / (Exp(REval) - Exp(-REval))

        'HCotan(X) = (Exp(X) + Exp(-X)) / (Exp(X) – Exp(-X))
        Case "hcotan": EvalBET = (Exp(REval) + Exp(-REval)) / (Exp(REval) - Exp(-REval))

        'HArcsin(X) = Log(X + Sqr(X * X + 1))
        Case "harcsin": EvalBET = Log(REval + Sqr(REval * REval + 1))

        'HArccos(X) = Log(X + Sqr(X * X – 1))
        Case "harccos": EvalBET = Log(REval + Sqr(REval * REval - 1))

        'HArctan(X) = Log((1 + X) / (1 – X)) / 2
        Case "harctan": EvalBET = Log((1 + REval) / (1 - REval)) / 2

        'HArcsec(X) = Log((Sqr(-X * X + 1) + 1) / X)
        Case "harcsec": EvalBET = Log((Sqr(-REval * REval + 1) + 1) / REval)

        'HArccosec(X) = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
        Case "harccosec": EvalBET = Log((Sgn(REval) * Sqr(REval * REval + 1) + 1) / REval)

        'HArccotan(X) = Log((X + 1) / (X – 1)) / 2
        Case "harccotan": EvalBET = Log((REval + 1) / (REval - 1)) / 2

        'user-defined variable
        Case Else: EvalBET = Variables.GetValue(Value)

    End Select
End Function

Public Function UsesIdentifier(IdName As String) As Boolean
    UsesIdentifier = UsesIdBET(IdName, 1)
End Function

Private Function UsesIdBET(IdName As String, EvTreeRoot As Long)
    If EvTree(EvTreeRoot).Value = IdName Then
        UsesIdBET = True
        Exit Function
    End If

    If EvTree(EvTreeRoot).Left > 0 Then
        If UsesIdBET(IdName, EvTree(EvTreeRoot).Left) Then
            UsesIdBET = True
            Exit Function
        End If
    End If

    If EvTree(EvTreeRoot).Right > 0 Then
        If UsesIdBET(IdName, EvTree(EvTreeRoot).Right) Then
            UsesIdBET = True
            Exit Function
        End If
    End If
End Function
