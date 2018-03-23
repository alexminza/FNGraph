VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDocumentProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Document Properties"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "MS Shell Dlg"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DocumentProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdValuesFont 
      Caption         =   "Font..."
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame fraAxes 
      Caption         =   "Axes"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdColor 
         Height          =   375
         Index           =   1
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboAxesStyle 
         Height          =   315
         ItemData        =   "DocumentProperties.frx":000C
         Left            =   960
         List            =   "DocumentProperties.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chkAxes 
         Caption         =   "&Axes"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   735
      End
      Begin VB.ComboBox cboAxesWidth 
         Height          =   315
         ItemData        =   "DocumentProperties.frx":003C
         Left            =   960
         List            =   "DocumentProperties.frx":0052
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblAxesColor 
         Caption         =   "Color:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblAxesStyle 
         Caption         =   "Style:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblAxesWidth 
         Caption         =   "Width:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Grid"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   2415
      Begin VB.CheckBox chkGrid 
         Caption         =   "&Grid"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdColor 
         Height          =   375
         Index           =   2
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboGridStyle 
         Height          =   315
         ItemData        =   "DocumentProperties.frx":0068
         Left            =   960
         List            =   "DocumentProperties.frx":0078
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboGridWidth 
         Height          =   315
         ItemData        =   "DocumentProperties.frx":0098
         Left            =   960
         List            =   "DocumentProperties.frx":00AE
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblGridColor 
         Caption         =   "Color:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblGridStyle 
         Caption         =   "Style:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblGridWidth 
         Caption         =   "Width:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   615
      End
   End
   Begin VB.Frame fraBackground 
      Caption         =   "&Background"
      Height          =   1695
      Left            =   2640
      TabIndex        =   21
      Top             =   1920
      Width           =   2415
      Begin VB.CommandButton cmdColor 
         Height          =   375
         Index           =   0
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblBackgroundColor 
         Caption         =   "Color:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraValues 
      Caption         =   "Values"
      Height          =   1695
      Left            =   2640
      TabIndex        =   16
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdColor 
         Height          =   375
         Index           =   3
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkValues 
         Caption         =   "&Values"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblValuesColor 
         Caption         =   "Color:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   120
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   25
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   24
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Defaults"
      Height          =   375
      Left            =   3960
      TabIndex        =   26
      Top             =   3840
      Width           =   1095
   End
End
Attribute VB_Name = "frmDocumentProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Private strValuesFontName As String, lngValuesFontSize As Long, blnValuesFontBold As Boolean, blnValuesFontItalic As Boolean

Private Sub Form_Load()
    cdlMain.flags = cdlCFScreenFonts + cdlCFForceFontExist

    cmdColor(0).BackColor = frmMain.ActiveForm.plngBackgroundColor
    cmdColor(1).BackColor = frmMain.ActiveForm.plngAxesColor
    cmdColor(2).BackColor = frmMain.ActiveForm.plngGridColor
    cmdColor(3).BackColor = frmMain.ActiveForm.plngValuesColor

    cboAxesStyle.ListIndex = frmMain.ActiveForm.plngAxesStyle
    cboAxesWidth.ListIndex = frmMain.ActiveForm.plngAxesWidth - 1
    cboGridStyle.ListIndex = frmMain.ActiveForm.plngGridStyle
    cboGridWidth.ListIndex = frmMain.ActiveForm.plngGridWidth - 1

    chkAxes.Value = IIf(frmMain.ActiveForm.pblnAxesVisible, vbChecked, vbUnchecked)
    chkGrid.Value = IIf(frmMain.ActiveForm.pblnGridVisible, vbChecked, vbUnchecked)
    chkValues.Value = IIf(frmMain.ActiveForm.pblnValuesVisible, vbChecked, vbUnchecked)

    strValuesFontName = frmMain.ActiveForm.pstrValuesFontName
    lngValuesFontSize = frmMain.ActiveForm.plngValuesFontSize
    blnValuesFontBold = frmMain.ActiveForm.pblnValuesFontBold
    blnValuesFontItalic = frmMain.ActiveForm.pblnValuesFontItalic
End Sub

Private Sub cmdValuesFont_Click()
    On Error GoTo ErrHandler

    cdlMain.FontName = strValuesFontName
    cdlMain.FontSize = lngValuesFontSize
    cdlMain.FontBold = blnValuesFontBold
    cdlMain.FontItalic = blnValuesFontItalic

    cdlMain.ShowFont

    strValuesFontName = cdlMain.FontName
    lngValuesFontSize = cdlMain.FontSize
    blnValuesFontBold = cdlMain.FontBold
    blnValuesFontItalic = cdlMain.FontItalic
    Exit Sub

ErrHandler:
    If Err.Number = cdlCancel Then Exit Sub
    ErrAssist vbOKOnly
End Sub

Private Sub cmdOK_Click()
    Dim blnNeedsRedraw As Boolean, blnTemp As Boolean

    'Background
    If frmMain.ActiveForm.plngBackgroundColor <> cmdColor(0).BackColor Then
        frmMain.ActiveForm.plngBackgroundColor = cmdColor(0).BackColor
        blnNeedsRedraw = True
    End If

    'Axes
    blnTemp = (chkAxes.Value = vbChecked)
    If frmMain.ActiveForm.pblnAxesVisible <> blnTemp Then
        frmMain.ActiveForm.pblnAxesVisible = blnTemp
        blnNeedsRedraw = True
    End If

    If frmMain.ActiveForm.plngAxesColor <> cmdColor(1).BackColor Then
        frmMain.ActiveForm.plngAxesColor = cmdColor(1).BackColor
        blnNeedsRedraw = True
    End If

    If frmMain.ActiveForm.plngAxesStyle <> cboAxesStyle.ListIndex Then
        frmMain.ActiveForm.plngAxesStyle = cboAxesStyle.ListIndex
        blnNeedsRedraw = True
    End If

    If frmMain.ActiveForm.plngAxesWidth <> cboAxesWidth.ListIndex + 1 Then
        frmMain.ActiveForm.plngAxesWidth = cboAxesWidth.ListIndex + 1
        blnNeedsRedraw = True
    End If

    'Grid
    blnTemp = (chkGrid.Value = vbChecked)
    If frmMain.ActiveForm.pblnGridVisible <> blnTemp Then
        frmMain.ActiveForm.pblnGridVisible = blnTemp
        blnNeedsRedraw = True
    End If

    If frmMain.ActiveForm.plngGridColor <> cmdColor(2).BackColor Then
        frmMain.ActiveForm.plngGridColor = cmdColor(2).BackColor
        blnNeedsRedraw = True
    End If

    If frmMain.ActiveForm.plngGridStyle <> cboGridStyle.ListIndex Then
        frmMain.ActiveForm.plngGridStyle = cboGridStyle.ListIndex
        blnNeedsRedraw = True
    End If

    If frmMain.ActiveForm.plngGridWidth <> cboGridWidth.ListIndex + 1 Then
        frmMain.ActiveForm.plngGridWidth = cboGridWidth.ListIndex + 1
        blnNeedsRedraw = True
    End If

    'Values
    blnTemp = (chkValues.Value = vbChecked)
    If frmMain.ActiveForm.pblnValuesVisible <> blnTemp Then
        frmMain.ActiveForm.pblnValuesVisible = blnTemp
        blnNeedsRedraw = True
    End If

    If frmMain.ActiveForm.plngValuesColor <> cmdColor(3).BackColor Then
        frmMain.ActiveForm.plngValuesColor = cmdColor(3).BackColor
        blnNeedsRedraw = True
    End If

    If frmMain.ActiveForm.pstrValuesFontName <> strValuesFontName Then
        frmMain.ActiveForm.pstrValuesFontName = strValuesFontName
        blnNeedsRedraw = True
    End If

    If frmMain.ActiveForm.plngValuesFontSize <> lngValuesFontSize Then
        frmMain.ActiveForm.plngValuesFontSize = lngValuesFontSize
        blnNeedsRedraw = True
    End If

    If frmMain.ActiveForm.pblnValuesFontBold <> blnValuesFontBold Then
        frmMain.ActiveForm.pblnValuesFontBold = blnValuesFontBold
        blnNeedsRedraw = True
    End If

    If frmMain.ActiveForm.pblnValuesFontItalic <> blnValuesFontItalic Then
        frmMain.ActiveForm.pblnValuesFontItalic = blnValuesFontItalic
        blnNeedsRedraw = True
    End If

    If blnNeedsRedraw Then
        frmMain.ActiveForm.FileChanged = True
        frmMain.ActiveForm.WindowRefresh
    End If
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDefaults_Click()
    cmdColor(0).BackColor = gBackgroundColor

    'Axes
    chkAxes.Value = vbChecked
    cmdColor(1).BackColor = gAxesColor
    cboAxesStyle.ListIndex = gAxesStyle
    cboAxesWidth.ListIndex = 0

    'Grid
    chkGrid.Value = vbChecked
    cmdColor(2).BackColor = gGridColor
    cboGridStyle.ListIndex = gGridStyle
    cboGridWidth.ListIndex = 0

    'Values
    chkValues.Value = vbChecked
    cmdColor(3).BackColor = gValuesColor
    strValuesFontName = gValuesFontName
    lngValuesFontSize = gValuesFontSize
    blnValuesFontBold = False
    blnValuesFontItalic = False
End Sub

Private Sub cmdColor_Click(Index As Integer)
    On Error GoTo ErrHandler

    cdlMain.Color = cmdColor(Index).BackColor
    cdlMain.ShowColor
    cmdColor(Index).BackColor = cdlMain.Color
    Exit Sub

ErrHandler:
    If Err.Number = cdlCancel Then Exit Sub
End Sub
