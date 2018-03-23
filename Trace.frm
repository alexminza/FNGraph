VERSION 5.00
Begin VB.Form frmTrace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trace"
   ClientHeight    =   2175
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
   Icon            =   "Trace.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraValues 
      Caption         =   "Values"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
      Begin VB.TextBox txtEval 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtEval 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   480
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblY 
         Caption         =   "&Y:"
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblX 
         Caption         =   "&X:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.HScrollBar hsbTrace 
      Height          =   255
      LargeChange     =   50
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   2
      Top             =   840
      Value           =   1
      Width           =   4095
   End
   Begin VB.ComboBox cboFunction 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   4095
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
Attribute VB_Name = "frmTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Dim llngLastPos As Long, llngLastIndex As Long
Dim X As Currency, Y As Currency, EvStat As Long, Pos As Long

Private Sub Form_Load()
    'Set window position
    If gblnSaveWindowsPos Then
        Me.Top = GetSetting(App.Title, gTraceRegKey, "Top", (Screen.Height - Me.Height) \ 2)
        Me.Left = GetSetting(App.Title, gTraceRegKey, "Left", (Screen.Width - Me.Width) \ 2)
    Else
        Me.Top = (Screen.Height - Me.Height) \ 2
        Me.Left = (Screen.Width - Me.Width) \ 2
    End If

    'Load graphs from the active document
    Dim I As Long
    For I = 1 To frmMain.ActiveForm.Graphs.Count
        cboFunction.AddItem frmMain.ActiveForm.Graphs(I).Expression.Expression
    Next I

    frmMain.ActiveForm.DrawWidth = 1
    frmMain.ActiveForm.DrawMode = vbInvert
    llngLastIndex = 0 'the form isn't unloaded completely!!! we always have to inititalise

    'Default to first item in list
    cboFunction.ListIndex = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub cboFunction_Click()
    RemoveTrace

    llngLastIndex = cboFunction.ListIndex + 1
    llngLastPos = 0

    hsbTrace.Max = frmMain.ActiveForm.Graphs(cboFunction.ListIndex + 1).ValuesCount
    If hsbTrace.Value = 1 Then
        hsbTrace_Change
    Else
        hsbTrace.Value = 1
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    RemoveTrace
    frmMain.ActiveForm.DrawMode = vbCopyPen

    'Save window position
    If gblnSaveWindowsPos Then
        SaveSetting App.Title, gTraceRegKey, "Top", Me.Top
        SaveSetting App.Title, gTraceRegKey, "Left", Me.Left
    End If
End Sub

Private Sub hsbTrace_Change()
    Pos = hsbTrace.Value
    frmMain.ActiveForm.Graphs(llngLastIndex).EvVal Pos, X, Y, EvStat
    txtEval(0).Text = CurToStr(X)

    Select Case EvStat
        Case EVAL_ERROR
            txtEval(1).Text = STR_ERROR

            If llngLastPos <> 0 Then
                frmMain.ActiveForm.GraphsTrace llngLastIndex, llngLastPos
                llngLastPos = 0
            End If
        Case Else
            txtEval(1).Text = CurToStr(Y)

            If llngLastPos <> 0 Then
                frmMain.ActiveForm.GraphsTrace llngLastIndex, llngLastPos
            End If
            frmMain.ActiveForm.GraphsTrace llngLastIndex, Pos
            llngLastPos = Pos
    End Select

    txtEval(0).Refresh
    txtEval(1).Refresh
End Sub

Private Sub hsbTrace_Scroll()
    hsbTrace_Change
End Sub

Private Sub txtEval_GotFocus(Index As Integer)
    txtEval(Index).SelStart = 0
    txtEval(Index).SelLength = Len(txtEval(Index).Text)
End Sub

Private Sub RemoveTrace()
    If llngLastIndex <> 0 Then
        If llngLastPos <> 0 Then
            frmMain.ActiveForm.GraphsTrace llngLastIndex, llngLastPos
        End If
    End If
End Sub
