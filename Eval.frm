VERSION 5.00
Begin VB.Form frmEval 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Evaluate"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "MS Shell Dlg"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Eval.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraValues 
      Caption         =   "Values"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4095
      Begin VB.TextBox txtEval 
         Height          =   285
         Index           =   0
         Left            =   480
         MaxLength       =   32
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtEval 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   32
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblX 
         Caption         =   "&X:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblY 
         Caption         =   "&Y:"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.ComboBox cboFunction 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "cboFunction"
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton cmdEval 
      Caption         =   "Evaluate"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblFunction 
      AutoSize        =   -1  'True
      Caption         =   "&Function:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmEval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Private EvExpression As New clsExpression

Private Sub Form_Load()
    'Set window position
    If gblnSaveWindowsPos Then
        Me.Top = GetSetting(App.Title, gEvalRegKey, "Top", (Screen.Height - Me.Height) \ 2)
        Me.Left = GetSetting(App.Title, gEvalRegKey, "Left", (Screen.Width - Me.Width) \ 2)
    Else
        Me.Top = (Screen.Height - Me.Height) \ 2
        Me.Left = (Screen.Width - Me.Width) \ 2
    End If

    'Get last used parameters
    cboFunction.Text = GetSetting(App.Title, gEvalRegKey, "Function", "")
    txtEval(0).Text = GetSetting(App.Title, gEvalRegKey, "X", CurToStr(0))

    'Load graphs from the active document
    If Not (frmMain.ActiveForm Is Nothing) Then
        Dim I As Long
        For I = 1 To frmMain.ActiveForm.Graphs.Count
            cboFunction.AddItem frmMain.ActiveForm.Graphs(I).Expression.Expression
        Next I
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Save window position
    If gblnSaveWindowsPos Then
        SaveSetting App.Title, gEvalRegKey, "Top", Me.Top
        SaveSetting App.Title, gEvalRegKey, "Left", Me.Left
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set EvExpression = Nothing
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEval_Click()
    On Error GoTo ErrHandler

    Dim X As Currency, Y As Currency, EvStat As Long

    If txtEval(0).Text = "" Then txtEval(0).Text = "0"
    X = StrToCur(txtEval(0).Text)
    txtEval(0).Text = CurToStr(X)

    If cboFunction.ListIndex = -1 Then
        'custom function specified
        cboFunction.Text = PrepareExpression(cboFunction.Text)
        If cboFunction.Text = "" Then
            MsgBox "Illegal Function value. Expression expected.", vbExclamation
            cboFunction.SetFocus
            Exit Sub
        Else
            'do we need to rebuild the expression BinEvTree?
            If EvExpression.Expression <> cboFunction.Text Then
                'if at least a document exists - hook to the active document's variables collection
                'else create a local empty collection
                Dim MyVariables As clsVariablesCollection
                If frmMain.ActiveForm Is Nothing Then
                    Set MyVariables = New clsVariablesCollection
                Else
                    Set MyVariables = frmMain.ActiveForm.Variables
                End If

                If CheckExpressionSyntax(cboFunction.Text, MyVariables) Then
                    cboFunction.SetFocus
                    Exit Sub
                Else
                    EvExpression.Expression = cboFunction.Text
                    Set EvExpression.Variables = MyVariables
                    EvExpression.Build
                End If
            End If
        End If
        Y = EvExpression.EvalFn(X, EvStat)
    Else
        'specified a function from the active document
        Y = frmMain.ActiveForm.Graphs(cboFunction.ListIndex + 1).Expression.EvalFn(X, EvStat)
    End If

    If EvStat = EVAL_ERROR Then
        txtEval(1).Text = STR_ERROR
    Else
        txtEval(1).Text = CurToStr(Y)
    End If

    SaveSettings

ProcExit:
    txtEval_GotFocus 0
    Exit Sub

ErrHandler:
    ErrAssist vbOKOnly
    txtEval(1).Text = ""
    GoTo ProcExit
End Sub

Private Sub cboFunction_Click()
    cmdEval_Click
End Sub

Private Sub txtEval_GotFocus(Index As Integer)
    txtEval(Index).SelStart = 0
    txtEval(Index).SelLength = Len(txtEval(Index).Text)
End Sub

Private Sub SaveSettings()
    SaveSetting App.Title, gEvalRegKey, "Function", cboFunction.Text
    SaveSetting App.Title, gEvalRegKey, "X", txtEval(0).Text
End Sub
