VERSION 5.00
Begin VB.Form frmIntegrate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integrate"
   ClientHeight    =   3495
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
   Icon            =   "Integrate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optPrecision 
      Caption         =   "&Precision:"
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton optSteps 
      Caption         =   "&Steps:"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtResult 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2520
      Width           =   3615
   End
   Begin VB.CommandButton cmdIntegrate 
      Caption         =   "Integrate"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame fraParameters 
      Caption         =   "Parameters"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
      Begin VB.TextBox txtGraph 
         Height          =   285
         Index           =   3
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtGraph 
         Height          =   285
         Index           =   2
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chkAbsValue 
         Caption         =   "By &absolute value (square)"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   2295
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
      Begin VB.TextBox txtGraph 
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   8
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblGraphMax 
         Caption         =   "Ma&x:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblGraphMin 
         Caption         =   "Mi&n:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.ComboBox cboFunction 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label lblResult 
      Caption         =   "&Result:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   615
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
Attribute VB_Name = "frmIntegrate"
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
        Me.Top = GetSetting(App.Title, gIntegrateRegKey, "Top", (Screen.Height - Me.Height) \ 2)
        Me.Left = GetSetting(App.Title, gIntegrateRegKey, "Left", (Screen.Width - Me.Width) \ 2)
    Else
        Me.Top = (Screen.Height - Me.Height) \ 2
        Me.Left = (Screen.Width - Me.Width) \ 2
    End If

    'Get last used parameters
    txtGraph(0).Text = GetSetting(App.Title, gIntegrateRegKey, "Min", CurToStr(gGraphMin))
    txtGraph(1).Text = GetSetting(App.Title, gIntegrateRegKey, "Max", CurToStr(gGraphMax))
    txtGraph(2).Text = GetSetting(App.Title, gIntegrateRegKey, "Precision", CurToStr(gGraphPrecision))
    txtGraph(3).Text = GetSetting(App.Title, gIntegrateRegKey, "Steps", LngToStr(gGraphSteps))
    If GetSetting(App.Title, gIntegrateRegKey, "PrecisionType", 1) = 1 Then
        optPrecision.Value = True
    Else
        optSteps.Value = True
    End If
    chkAbsValue.Value = IIf(GetSetting(App.Title, gIntegrateRegKey, "ByAbsValue", True), vbChecked, vbUnchecked)

    'Load graphs from the active document
    Dim I As Long
    For I = 1 To frmMain.ActiveForm.Graphs.Count
        cboFunction.AddItem frmMain.ActiveForm.Graphs(I).Expression.Expression
    Next I

    'Default to first item in list
    cboFunction.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Save window position
    If gblnSaveWindowsPos Then
        SaveSetting App.Title, gIntegrateRegKey, "Top", Me.Top
        SaveSetting App.Title, gIntegrateRegKey, "Left", Me.Left
    End If
End Sub

Private Sub cmdIntegrate_Click()
    On Error GoTo ErrHandler
    Dim curMin As Currency, curMax As Currency, curPrecision As Currency
    Dim lngSteps As Long

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
    Integrate cboFunction.ListIndex + 1, curMin, curMax, curPrecision, (chkAbsValue.Value = vbChecked)

ProcExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHandler:
    ErrAssist vbOKOnly
    GoTo ProcExit
End Sub

Private Sub Integrate(GraphIndex As Long, IMA As Currency, IMB As Currency, IP As Currency, ByAbsVal As Boolean)
    Dim ICX As Currency, ICY As Currency, S As Currency, EvStat As Long

    ICY = frmMain.ActiveForm.Graphs(GraphIndex).Expression.EvalFn(IMA, EvStat)
    If EvStat = EVAL_ERROR Then
        txtResult.Text = STR_ERROR
        Exit Sub
    End If
    If ByAbsVal Then
        S = Abs(ICY)
    Else
        S = ICY
    End If

    ICY = frmMain.ActiveForm.Graphs(GraphIndex).Expression.EvalFn(IMB, EvStat)
    If EvStat = EVAL_ERROR Then
        txtResult.Text = STR_ERROR
        Exit Sub
    End If
    If ByAbsVal Then
        S = S + Abs(ICY)
    Else
        S = S + ICY
    End If

    ICX = IMA + IP
    Do While ICX < IMB
        ICY = frmMain.ActiveForm.Graphs(GraphIndex).Expression.EvalFn(ICX, EvStat)
        If EvStat = 1 Then 'EVAL_ERROR
            txtResult.Text = STR_ERROR
            Exit Sub
        End If
        If ByAbsVal Then
            S = S + 2 * Abs(ICY)
        Else
            S = S + 2 * ICY
        End If
        ICX = ICX + IP
    Loop

    S = S * IP / 2
    txtResult.Text = CurToStr(S)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cboFunction_Click()
    txtResult.Text = ""
End Sub

Private Sub SaveSettings()
    SaveSetting App.Title, gIntegrateRegKey, "Min", txtGraph(0).Text
    SaveSetting App.Title, gIntegrateRegKey, "Max", txtGraph(1).Text
    SaveSetting App.Title, gIntegrateRegKey, "Precision", txtGraph(2).Text
    SaveSetting App.Title, gIntegrateRegKey, "Steps", txtGraph(3).Text
    SaveSetting App.Title, gIntegrateRegKey, "PrecisionType", IIf(optPrecision.Value, 1, 2)
    SaveSetting App.Title, gIntegrateRegKey, "ByAbsValue", (chkAbsValue.Value = vbChecked)
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

Private Sub txtResult_GotFocus()
    txtResult.SelStart = 0
    txtResult.SelLength = Len(txtResult.Text)
End Sub
