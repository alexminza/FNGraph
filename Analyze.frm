VERSION 5.00
Begin VB.Form frmAnalyze 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analyze"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "MS Shell Dlg"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Analyze.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optPrecision 
      Caption         =   "&Precision:"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton optSteps 
      Caption         =   "&Steps:"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopySel 
      Caption         =   "Copy s&el"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ListBox lstResults 
      Height          =   1230
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   11
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame fraParameters 
      Caption         =   "Parameters"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
      Begin VB.TextBox txtGraph 
         Height          =   285
         Index           =   3
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtGraph 
         Height          =   285
         Index           =   2
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtGraph 
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   8
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtGraph 
         Height          =   285
         Index           =   1
         Left            =   960
         MaxLength       =   8
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblGraphMin 
         Caption         =   "Mi&n:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblGraphMax 
         Caption         =   "Ma&x:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "Analyze"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ComboBox cboFunction 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   4335
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
Attribute VB_Name = "frmAnalyze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Private Sub Form_Load()
    'Set window position
    If gblnSaveWindowsPos Then
        Me.Top = GetSetting(App.Title, gAnalyzeRegKey, "Top", (Screen.Height - Me.Height) \ 2)
        Me.Left = GetSetting(App.Title, gAnalyzeRegKey, "Left", (Screen.Width - Me.Width) \ 2)
    Else
        Me.Top = (Screen.Height - Me.Height) \ 2
        Me.Left = (Screen.Width - Me.Width) \ 2
    End If

    'Get last used parameters
    txtGraph(0).Text = GetSetting(App.Title, gAnalyzeRegKey, "Min", CurToStr(gGraphMin))
    txtGraph(1).Text = GetSetting(App.Title, gAnalyzeRegKey, "Max", CurToStr(gGraphMax))
    txtGraph(2).Text = GetSetting(App.Title, gAnalyzeRegKey, "Precision", CurToStr(gGraphPrecision))
    txtGraph(3).Text = GetSetting(App.Title, gAnalyzeRegKey, "Steps", LngToStr(gGraphSteps))
    If GetSetting(App.Title, gAnalyzeRegKey, "PrecisionType", 1) = 1 Then
        optPrecision.Value = True
    Else
        optSteps.Value = True
    End If

    'Load graphs from the active document
    Dim I As Long
    For I = 1 To frmMain.ActiveForm.Graphs.Count
        cboFunction.AddItem frmMain.ActiveForm.Graphs(I).Expression.Expression
    Next I

    cboFunction.ListIndex = 0
    UpdateCopyCmds
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Save window position
    If gblnSaveWindowsPos Then
        SaveSetting App.Title, gAnalyzeRegKey, "Top", Me.Top
        SaveSetting App.Title, gAnalyzeRegKey, "Left", Me.Left
    End If
End Sub

Private Sub cboFunction_Click()
    lstResults.Clear
    UpdateCopyCmds
End Sub

Private Sub cmdAnalyze_Click()
    On Error GoTo ErrHandler
    Dim GraphIndex As Long
    Dim curMin As Currency, curMax As Currency, curPrecision As Currency, lngSteps As Long

    'Validate Min and Max
    curMin = StrToCur(txtGraph(0).Text)
    curMax = StrToCur(txtGraph(1).Text)
    If curMin >= curMax Then
        MsgBox "Invalid Min or Max value. Max must be greater than Min.", vbExclamation
        txtGraph_GotFocus 0
        txtGraph(0).SetFocus
        Exit Sub
    End If
    txtGraph(0).Text = CurToStr(curMin)
    txtGraph(1).Text = CurToStr(curMax)

    'Validate Precision
    If optPrecision.Value Then
        curPrecision = StrToCur(txtGraph(2).Text)
        If (curPrecision <= 0) Or (curPrecision > (curMax - curMin)) Then
            MsgBox "Invalid Precision value. Must be greater than 0 and less than the chosen interval.", vbExclamation
            txtGraph_GotFocus 2
            txtGraph(2).SetFocus
            Exit Sub
        End If
        txtGraph(2).Text = CurToStr(curPrecision)
    Else
        lngSteps = StrToLng(txtGraph(3).Text)
        If (lngSteps < 1) Or (lngSteps > 10000) Then
            MsgBox "Invalid Steps value. Must be between 1 and 10000.", vbExclamation
            txtGraph_GotFocus 3
            txtGraph(3).SetFocus
            Exit Sub
        End If
        txtGraph(3).Text = LngToStr(lngSteps)

        curPrecision = (curMax - curMin) / lngSteps
        If curPrecision = 0 Then
            MsgBox "The resulting precision is 0. Choose less steps.", vbExclamation
            txtGraph_GotFocus 3
            txtGraph(3).SetFocus
            Exit Sub
        End If
    End If

    Screen.MousePointer = vbHourglass
    SaveSettings

    lstResults.Clear
    lstResults.Visible = False
    GraphIndex = cboFunction.ListIndex + 1

    AnalyzeFn GraphIndex, curMin, curMax, curPrecision
    If lstResults.ListCount = 0 Then lstResults.AddItem "No zeroes found"

    lstResults.Visible = True
    UpdateCopyCmds

ProcExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHandler:
    ErrAssist vbOKOnly
    GoTo ProcExit
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    Screen.MousePointer = vbHourglass
    Dim I As Long, Buffer As String

    Buffer = "Function: " & cboFunction.Text & vbCrLf & "Min: " & _
    txtGraph(0).Text & vbCrLf & "Max: " & txtGraph(1).Text & vbCrLf

    If optPrecision.Value Then
        Buffer = Buffer & "Precision: " & txtGraph(2).Text
    Else
        Buffer = Buffer & "Steps: " & txtGraph(3).Text
    End If
    Buffer = Buffer & vbCrLf & vbCrLf & "Zeroes:"

    For I = 0 To lstResults.ListCount - 1
        Buffer = Buffer & vbCrLf & lstResults.List(I)
    Next I

    Clipboard.Clear
    Clipboard.SetText Buffer
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCopySel_Click()
    Screen.MousePointer = vbHourglass
    Dim I As Long, Buffer As String

    If lstResults.SelCount > 0 Then
        For I = 0 To lstResults.ListCount - 1
            If lstResults.Selected(I) Then
                Buffer = lstResults.List(I)
                Exit For
            End If
        Next I

        If lstResults.SelCount > 1 Then
            For I = I + 1 To lstResults.ListCount - 1
                If lstResults.Selected(I) Then
                    Buffer = Buffer & vbCrLf & lstResults.List(I)
                End If
            Next I
        End If
    End If

    Clipboard.Clear
    Clipboard.SetText Buffer
    Screen.MousePointer = vbDefault
End Sub

Private Sub lstResults_Click()
    cmdCopySel.Enabled = (lstResults.SelCount > 0)
End Sub

Private Sub optPrecision_Click()
    txtGraph(2).Enabled = True
    txtGraph(2).BackColor = vbWindowBackground
    txtGraph(3).Enabled = False
    txtGraph(3).BackColor = vbButtonFace
End Sub

Private Sub optSteps_Click()
    txtGraph(2).Enabled = False
    txtGraph(2).BackColor = vbButtonFace
    txtGraph(3).Enabled = True
    txtGraph(3).BackColor = vbWindowBackground
End Sub

Private Sub txtGraph_GotFocus(Index As Integer)
    txtGraph(Index).SelStart = 0
    txtGraph(Index).SelLength = Len(txtGraph(Index).Text)
End Sub

Private Sub UpdateCopyCmds()
    cmdCopy.Enabled = (lstResults.ListCount > 0)
    cmdCopySel.Enabled = False
End Sub

Private Sub SaveSettings()
    SaveSetting App.Title, gAnalyzeRegKey, "Min", txtGraph(0).Text
    SaveSetting App.Title, gAnalyzeRegKey, "Max", txtGraph(1).Text
    SaveSetting App.Title, gAnalyzeRegKey, "Precision", txtGraph(2).Text
    SaveSetting App.Title, gAnalyzeRegKey, "Steps", txtGraph(3).Text
    SaveSetting App.Title, gAnalyzeRegKey, "PrecisionType", IIf(optPrecision.Value, 1, 2)
End Sub

Private Sub AnalyzeFn(GraphIndex As Long, AMA As Currency, AMB As Currency, AP As Currency)
    Dim AXA As Currency, AXB As Currency, AX As Currency, AYA As Currency, AYB As Currency
    Dim EvStat As Long, MaxGap As Long

    AXA = AMA
    MaxGap = frmMain.ActiveForm.Graphs(GraphIndex).MaxGap

    Do While AXA < AMB
        AXB = AXA + AP
        AYA = frmMain.ActiveForm.Graphs(GraphIndex).Expression.EvalFn(AXA, EvStat)
        If EvStat <> 1 Then 'EVAL_ERROR
            If AYA = 0 Then
                lstResults.AddItem CurToStr(AXA)
            Else
                AYB = frmMain.ActiveForm.Graphs(GraphIndex).Expression.EvalFn(AXB, EvStat)
                If EvStat <> 1 Then 'EVAL_ERROR
                    'Do not add the same zero twice
                    If AYB <> 0 Then
                        'Does the (AXA,AYA)-(AXB,AYB) segment intersect the Ox axis?
                        If Sgn(AYA) <> Sgn(AYB) Then
                            'Is it a discontinuance?
                            If Abs(AYA - AYB) < MaxGap Then
                                If BisectFn(GraphIndex, AXA, AXB, AX) Then lstResults.AddItem CurToStr(AX)
                            End If
                        End If
                    End If
                End If
            End If
        End If

        AXA = AXB
    Loop
End Sub

Private Function BisectFn(GraphIndex As Long, BMA As Currency, BMB As Currency, BX As Currency) As Boolean
    Dim BCX As Currency, BCY As Currency, EvStat As Long

    Do
        BCX = (BMA + BMB) / 2
        BCY = frmMain.ActiveForm.Graphs(GraphIndex).Expression.EvalFn(BCX, EvStat)
        If (BCY = 0) Or (Abs(BMB - BMA) <= 0.0001) Then
            If EvStat <> 1 Then 'EVAL_ERROR
                BX = BCX
                BisectFn = True
                Exit Function
            Else
                BisectFn = False
                Exit Function
            End If
        Else
            If Sgn(frmMain.ActiveForm.Graphs(GraphIndex).Expression.EvalFn(BMA, 0)) <> Sgn(BCY) Then
                BMB = BCX
            Else
                BMA = BCX
            End If
        End If
    Loop
End Function
