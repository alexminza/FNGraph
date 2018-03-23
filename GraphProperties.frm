VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGraphProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Graph Properties"
   ClientHeight    =   4815
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
   Icon            =   "GraphProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGraph 
      Height          =   285
      Index           =   4
      Left            =   1080
      MaxLength       =   255
      TabIndex        =   22
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Frame fraAppearance 
      Caption         =   "Appearance"
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   4095
      Begin VB.CommandButton cmdGraphColor 
         Height          =   255
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1200
         Width           =   975
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "&Visible"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox chkDrawLines 
         Caption         =   "&Lines"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkDrawPoints 
         Caption         =   "Poin&ts"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cboLinesWidth 
         Height          =   315
         ItemData        =   "GraphProperties.frx":000C
         Left            =   1920
         List            =   "GraphProperties.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cboPointsSize 
         Height          =   315
         ItemData        =   "GraphProperties.frx":0038
         Left            =   1920
         List            =   "GraphProperties.frx":004E
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblGraphColor 
         Caption         =   "&Color:"
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblLinesWidth 
         Caption         =   "&Width:"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblPointsSize 
         Caption         =   "&Size:"
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   120
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraProperties 
      Caption         =   "Properties"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4095
      Begin VB.TextBox txtGraph 
         Height          =   285
         Index           =   3
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtGraph 
         Height          =   285
         Index           =   2
         Left            =   3000
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
      Begin VB.Label lblMaxGap 
         Caption         =   "Max &gap:"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblGraphPrecision 
         Caption         =   "&Precision:"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   360
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
      Begin VB.Label lblGraphMax 
         Caption         =   "Ma&x:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   24
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      Top             =   4320
      Width           =   1095
   End
   Begin VB.ComboBox cboFunction 
      Height          =   315
      ItemData        =   "GraphProperties.frx":0064
      Left            =   120
      List            =   "GraphProperties.frx":0066
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label lblGraphDescription 
      AutoSize        =   -1  'True
      Caption         =   "&Description:"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   855
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
Attribute VB_Name = "frmGraphProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Public AddNewGraph As Boolean, CurrentGraph As clsGraph

Private Sub Form_Load()
    'Set window position
    If gblnSaveWindowsPos Then
        Me.Top = GetSetting(App.Title, gGraphPropertiesRegKey, "Top", (Screen.Height - Me.Height) \ 2)
        Me.Left = GetSetting(App.Title, gGraphPropertiesRegKey, "Left", (Screen.Width - Me.Width) \ 2)
    Else
        Me.Top = (Screen.Height - Me.Height) \ 2
        Me.Left = (Screen.Width - Me.Width) \ 2
    End If

    cdlMain.flags = cdlCCRGBInit

    'Load recent functions
    Dim I As Long, S As String
    For I = 1 To gRecentFunctionsCount
        S = GetSetting(App.Title, gRecentFunctionsRegKey, LngToStr(I), "")
        If S <> "" Then cboFunction.AddItem S
    Next I

    'Add or Properties
    If AddNewGraph Then
        Me.Caption = "Add Graph"
    Else
        Me.Caption = "Graph Properties"
    End If

    cboFunction.Text = CurrentGraph.Expression.Expression
    txtGraph(0).Text = CurToStr(CurrentGraph.Min)
    txtGraph(1).Text = CurToStr(CurrentGraph.Max)
    txtGraph(2).Text = CurToStr(CurrentGraph.Precision)
    txtGraph(3).Text = CurToStr(CurrentGraph.MaxGap)
    chkDrawLines.Value = IIf(CurrentGraph.DrawLines, vbChecked, vbUnchecked)
    chkDrawPoints.Value = IIf(CurrentGraph.DrawPoints, vbChecked, vbUnchecked)
    cboLinesWidth.ListIndex = CurrentGraph.LinesWidth - 1
    cboPointsSize.ListIndex = CurrentGraph.PointsSize - 1
    chkVisible.Value = IIf(CurrentGraph.Visible, vbChecked, vbUnchecked)
    cmdGraphColor.BackColor = CurrentGraph.Color
    txtGraph(4).Text = CurrentGraph.Description
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Save window position
    If gblnSaveWindowsPos Then
        SaveSetting App.Title, gGraphPropertiesRegKey, "Top", Me.Top
        SaveSetting App.Title, gGraphPropertiesRegKey, "Left", Me.Left
    End If
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandler

    Dim blnNeedsBuild As Boolean, blnNeedsEval As Boolean, blnNeedsRedraw As Boolean, blnPropsChanged As Boolean
    Dim curTmp1 As Currency, curTmp2 As Currency, lngTmp As Long, blnTmp As Boolean

    'Validate the function string
    cboFunction.Text = PrepareExpression(cboFunction.Text)
    If cboFunction.Text = "" Then
        MsgBox "Illegal Function value. Expression expected.", vbExclamation
        cboFunction.SetFocus
        Exit Sub
    Else
        If CurrentGraph.Expression.Expression <> cboFunction.Text Then
            If CheckExpressionSyntax(cboFunction.Text, frmMain.ActiveForm.Variables) Then
                cboFunction.SetFocus
                Exit Sub
            Else
                'Set the Expression property
                CurrentGraph.Expression.Expression = cboFunction.Text
                SaveRecentFunctions
    
                blnNeedsBuild = True
                blnNeedsEval = True
                blnNeedsRedraw = True
                blnPropsChanged = True
            End If
        End If
    End If

    'Validate Min and Max
    curTmp1 = StrToCur(txtGraph(0).Text)
    curTmp2 = StrToCur(txtGraph(1).Text)
    If (curTmp1 < -10000) Or (curTmp1 > 10000) Then
        MsgBox "Invalid Min value. Must be between -10000 and 10000.", vbExclamation
        txtGraph_GotFocus 0
        txtGraph(0).SetFocus
        Exit Sub
    End If
    If (curTmp2 < -10000) Or (curTmp2 > 10000) Then
        MsgBox "Invalid Max value. Must be between -10000 and 10000.", vbExclamation
        txtGraph_GotFocus 1
        txtGraph(1).SetFocus
        Exit Sub
    End If
    If curTmp1 >= curTmp2 Then
        MsgBox "Invalid Min or Max value. Max must be greater than Min.", vbExclamation
        txtGraph_GotFocus 0
        txtGraph(0).SetFocus
        Exit Sub
    End If
    If (CurrentGraph.Min <> curTmp1) Or (CurrentGraph.Max <> curTmp2) Then
        CurrentGraph.Min = curTmp1
        CurrentGraph.Max = curTmp2
        blnNeedsEval = True
        blnNeedsRedraw = True
        blnPropsChanged = True
    End If

    'Validate Precision
    curTmp1 = StrToCur(txtGraph(2).Text)
    If (curTmp1 <= 0) Or (curTmp1 > 1) Then
        MsgBox "Invalid Precision value. Must be between 0 and 1.", vbExclamation
        txtGraph_GotFocus 2
        txtGraph(2).SetFocus
        Exit Sub
    End If
    If CurrentGraph.Precision <> curTmp1 Then
        CurrentGraph.Precision = curTmp1
        blnNeedsEval = True
        blnNeedsRedraw = True
        blnPropsChanged = True
    End If

    'Validate Max gap
    lngTmp = CLng(txtGraph(3).Text)
    If (lngTmp < 1) Or (lngTmp > 500) Then
        MsgBox "Invalid Max gap value. Must be between 1 and 500.", vbExclamation
        txtGraph_GotFocus 3
        txtGraph(3).SetFocus
        Exit Sub
    End If
    If CurrentGraph.MaxGap <> lngTmp Then
        CurrentGraph.MaxGap = lngTmp
        blnNeedsRedraw = True
        blnPropsChanged = True
    End If

    'Set DrawLines
    blnTmp = (chkDrawLines.Value = vbChecked)
    If CurrentGraph.DrawLines <> blnTmp Then
        CurrentGraph.DrawLines = blnTmp
        blnNeedsRedraw = True
        blnPropsChanged = True
    End If

    'Set DrawPoints
    blnTmp = (chkDrawPoints.Value = vbChecked)
    If CurrentGraph.DrawPoints <> blnTmp Then
        CurrentGraph.DrawPoints = blnTmp
        blnNeedsRedraw = True
        blnPropsChanged = True
    End If

    'Set LinesWidth
    lngTmp = cboLinesWidth.ListIndex + 1
    If CurrentGraph.LinesWidth <> lngTmp Then
        CurrentGraph.LinesWidth = lngTmp
        blnNeedsRedraw = True
        blnPropsChanged = True
    End If

    'Set PointsSize
    lngTmp = cboPointsSize.ListIndex + 1
    If CurrentGraph.PointsSize <> lngTmp Then
        CurrentGraph.PointsSize = lngTmp
        blnNeedsRedraw = True
        blnPropsChanged = True
    End If

    'Set Visible
    blnTmp = (chkVisible.Value = vbChecked)
    If CurrentGraph.Visible <> blnTmp Then
        CurrentGraph.Visible = blnTmp
        blnNeedsRedraw = True
        blnPropsChanged = True
    End If

    'Set Color
    lngTmp = cmdGraphColor.BackColor
    If CurrentGraph.Color <> lngTmp Then
        CurrentGraph.Color = lngTmp
        blnNeedsRedraw = True
        blnPropsChanged = True
    End If

    'Set Description
    If txtGraph(4) <> CurrentGraph.Description Then
        CurrentGraph.Description = txtGraph(4).Text
        blnPropsChanged = True
    End If


    frmGraphs.NeedsBuild = blnNeedsBuild
    frmGraphs.NeedsEval = blnNeedsEval
    frmGraphs.NeedsRedraw = blnNeedsRedraw
    frmGraphs.PropsChanged = blnPropsChanged

    Unload Me
    Exit Sub

ErrHandler:
    ErrAssist vbOKOnly
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGraphColor_Click()
    On Error GoTo ErrHandler

    cdlMain.Color = cmdGraphColor.BackColor
    cdlMain.ShowColor
    cmdGraphColor.BackColor = cdlMain.Color
    Exit Sub

ErrHandler:
    If Err.Number = cdlCancel Then Exit Sub
End Sub

Private Sub txtGraph_GotFocus(Index As Integer)
    txtGraph(Index).SelStart = 0
    txtGraph(Index).SelLength = Len(txtGraph(Index).Text)
End Sub

Private Sub SaveRecentFunctions()
    Dim I As Long, T As Long

    'save current function
    SaveSetting App.Title, gRecentFunctionsRegKey, "1", cboFunction.Text

    T = 1
    For I = 0 To cboFunction.ListCount - 1
        If cboFunction.Text <> cboFunction.List(I) Then
            T = T + 1
            SaveSetting App.Title, gRecentFunctionsRegKey, LngToStr(T), cboFunction.List(I)

            If T = gRecentFunctionsCount Then Exit For
        End If
    Next I
End Sub
