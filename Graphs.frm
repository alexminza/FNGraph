VERSION 5.00
Begin VB.Form frmGraphs 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Graphs"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "MS Shell Dlg"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Graphs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "&Properties"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox lstFunctions 
      Height          =   2145
      IntegralHeight  =   0   'False
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmGraphs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Public PropsChanged As Boolean, NeedsBuild As Boolean, NeedsEval As Boolean, NeedsRedraw As Boolean
Private lngCurrentGraph As Long, blnUpdatingSelection As Boolean

Private Sub Form_Load()
    frmMain.IsGraphsWindowLoaded = True

    'Set window position
    If gblnSaveWindowsPos Then
        Dim lpwndpl As WINDOWPLACEMENT

        lpwndpl.length = Len(lpwndpl)
        lpwndpl.flags = 0
        lpwndpl.showCmd = SW_SHOWNORMAL
        lpwndpl.rcNormalPosition.Left = GetSetting(App.Title, gGraphsRegKey, "rcNormalPosition.left", (Screen.Width - Me.Width) \ Screen.TwipsPerPixelX \ 2)
        lpwndpl.rcNormalPosition.Top = GetSetting(App.Title, gGraphsRegKey, "rcNormalPosition.top", (Screen.Height - Me.Height) \ Screen.TwipsPerPixelY \ 2)
        lpwndpl.rcNormalPosition.Right = GetSetting(App.Title, gGraphsRegKey, "rcNormalPosition.right", lpwndpl.rcNormalPosition.Left + Me.Width \ Screen.TwipsPerPixelX)
        lpwndpl.rcNormalPosition.Bottom = GetSetting(App.Title, gGraphsRegKey, "rcNormalPosition.bottom", lpwndpl.rcNormalPosition.Top + Me.Height \ Screen.TwipsPerPixelY)

        SetWindowPlacement Me.hWnd, lpwndpl
    Else
        Me.Top = (Screen.Height - Me.Height) \ 2
        Me.Left = (Screen.Width - Me.Width) \ 2
    End If

    UpdateList
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Save window position
    If gblnSaveWindowsPos Then
        Dim lpwndpl As WINDOWPLACEMENT

        lpwndpl.length = Len(lpwndpl)
        GetWindowPlacement Me.hWnd, lpwndpl
        SaveSetting App.Title, gGraphsRegKey, "rcNormalPosition.left", lpwndpl.rcNormalPosition.Left
        SaveSetting App.Title, gGraphsRegKey, "rcNormalPosition.top", lpwndpl.rcNormalPosition.Top
        SaveSetting App.Title, gGraphsRegKey, "rcNormalPosition.right", lpwndpl.rcNormalPosition.Right
        SaveSetting App.Title, gGraphsRegKey, "rcNormalPosition.bottom", lpwndpl.rcNormalPosition.Bottom
    End If

    frmMain.IsGraphsWindowLoaded = False
    frmMain.SetFocus 'VB5 Bug
End Sub

Private Sub Form_Resize()
    Dim CmdLeft As Long
    CmdLeft = Me.ScaleWidth - 1215

    If CmdLeft > 1500 Then
        lstFunctions.Width = CmdLeft - 240

        cmdAdd.Left = CmdLeft
        cmdProperties.Left = CmdLeft
        cmdDelete.Left = CmdLeft
    End If
    If Me.ScaleHeight > 1575 Then
        lstFunctions.Height = Me.ScaleHeight - 240
    End If

    DoEvents
End Sub

Private Sub ApplyChanges()
    If NeedsBuild Then
        frmMain.ActiveForm.Graphs(lngCurrentGraph).Expression.Build
    End If

    If NeedsEval Then
        frmMain.ActiveForm.Graphs(lngCurrentGraph).CacheValues
    End If

    If NeedsRedraw Then
        frmMain.ActiveForm.WindowRefresh
    End If

    frmMain.ActiveForm.FileChanged = True
End Sub

Private Sub cmdAdd_Click()
    frmMain.ActiveForm.GraphsAdd
    lngCurrentGraph = frmMain.ActiveForm.Graphs.Count

    Set frmGraphProperties.CurrentGraph = frmMain.ActiveForm.Graphs(lngCurrentGraph)
    frmGraphProperties.AddNewGraph = True
    PropsChanged = False
    frmGraphProperties.Show vbModal, frmMain
    DoEvents

    Screen.MousePointer = vbHourglass
    If PropsChanged Then
        ApplyChanges

        blnUpdatingSelection = True
        lstFunctions.AddItem frmMain.ActiveForm.Graphs(lngCurrentGraph).DisplayName
        lstFunctions.Selected(lngCurrentGraph - 1) = frmMain.ActiveForm.Graphs(lngCurrentGraph).Visible
        blnUpdatingSelection = False
        lstFunctions.ListIndex = lngCurrentGraph - 1

        UpdateCommands
    Else
        frmMain.ActiveForm.Graphs.Remove lngCurrentGraph
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdProperties_Click()
    lngCurrentGraph = lstFunctions.ListIndex + 1

    Set frmGraphProperties.CurrentGraph = frmMain.ActiveForm.Graphs(lngCurrentGraph)
    frmGraphProperties.AddNewGraph = False
    PropsChanged = False
    frmGraphProperties.Show vbModal, frmMain
    DoEvents

    Screen.MousePointer = vbHourglass
    If PropsChanged Then
        ApplyChanges

        blnUpdatingSelection = True
        lstFunctions.List(lstFunctions.ListIndex) = frmMain.ActiveForm.Graphs(lngCurrentGraph).DisplayName
        lstFunctions.Selected(lstFunctions.ListIndex) = frmMain.ActiveForm.Graphs(lngCurrentGraph).Visible
        blnUpdatingSelection = False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDelete_Click()
    'confirmation
    If gblnConfirmDelete Then
        If MsgBox("Are you sure you want to delete this graph?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    Screen.MousePointer = vbHourglass

    Dim LastItem As Long
    LastItem = lstFunctions.ListIndex
    frmMain.ActiveForm.Graphs.Remove LastItem + 1
    lstFunctions.RemoveItem LastItem

    If LastItem < lstFunctions.ListCount Then
        lstFunctions.ListIndex = LastItem
    Else
        lstFunctions.ListIndex = LastItem - 1
    End If

    frmMain.ActiveForm.FileChanged = True
    frmMain.ActiveForm.WindowRefresh
    UpdateCommands

    Screen.MousePointer = vbDefault
End Sub

Private Sub lstFunctions_ItemCheck(Item As Integer)
    If Not blnUpdatingSelection Then
        If lstFunctions.ListCount > 0 Then
            DoEvents
            frmMain.ActiveForm.Graphs(Item + 1).Visible = lstFunctions.Selected(Item)
            frmMain.ActiveForm.FileChanged = True
            frmMain.ActiveForm.WindowRefresh
        End If
    End If
End Sub

Private Sub UpdateCommands()
    If frmMain.ActiveForm Is Nothing Then
        cmdAdd.Enabled = False
        cmdProperties.Enabled = False
        cmdDelete.Enabled = False
    Else
        cmdAdd.Enabled = True 'Unlimited number of graphs
        cmdProperties.Enabled = (lstFunctions.ListCount > 0)
        cmdDelete.Enabled = cmdProperties.Enabled
    End If

    frmMain.UpdateGraphCmds
End Sub

Public Sub UpdateList()
    If frmMain.ActiveForm Is Nothing Then
        Me.Caption = "Graphs"
        lstFunctions.Clear
        UpdateCommands
        Exit Sub
    End If

    Me.Caption = "Graphs - " & frmMain.ActiveForm.Caption

    Dim I As Long
    lstFunctions.Clear
    blnUpdatingSelection = True
    For I = 1 To frmMain.ActiveForm.Graphs.Count
        lstFunctions.AddItem frmMain.ActiveForm.Graphs(I).DisplayName
        lstFunctions.Selected(I - 1) = frmMain.ActiveForm.Graphs(I).Visible
    Next I
    blnUpdatingSelection = False
    If lstFunctions.ListCount > 0 Then lstFunctions.ListIndex = 0

    UpdateCommands
End Sub
