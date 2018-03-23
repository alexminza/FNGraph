VERSION 5.00
Begin VB.Form frmVariables 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Variables"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   330
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstVariables 
      Height          =   2145
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "&Properties"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Public PropsChanged As Boolean

Private Sub Form_Load()
    frmMain.IsVariablesWindowLoaded = True

    'Set window position
    If gblnSaveWindowsPos Then
        Dim lpwndpl As WINDOWPLACEMENT

        lpwndpl.length = Len(lpwndpl)
        lpwndpl.flags = 0
        lpwndpl.showCmd = SW_SHOWNORMAL
        lpwndpl.rcNormalPosition.Left = GetSetting(App.Title, gVariablesRegKey, "rcNormalPosition.left", (Screen.Width - Me.Width) \ Screen.TwipsPerPixelX \ 2 + 30)
        lpwndpl.rcNormalPosition.Top = GetSetting(App.Title, gVariablesRegKey, "rcNormalPosition.top", (Screen.Height - Me.Height) \ Screen.TwipsPerPixelY \ 2 + 30)
        lpwndpl.rcNormalPosition.Right = GetSetting(App.Title, gVariablesRegKey, "rcNormalPosition.right", lpwndpl.rcNormalPosition.Left + Me.Width \ Screen.TwipsPerPixelX)
        lpwndpl.rcNormalPosition.Bottom = GetSetting(App.Title, gVariablesRegKey, "rcNormalPosition.bottom", lpwndpl.rcNormalPosition.Top + Me.Height \ Screen.TwipsPerPixelY)

        SetWindowPlacement Me.hWnd, lpwndpl
    Else
        Me.Top = (Screen.Height - Me.Height) \ 2 + 450
        Me.Left = (Screen.Width - Me.Width) \ 2 + 450
    End If

    UpdateList
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Save window position
    If gblnSaveWindowsPos Then
        Dim lpwndpl As WINDOWPLACEMENT

        lpwndpl.length = Len(lpwndpl)
        GetWindowPlacement Me.hWnd, lpwndpl
        SaveSetting App.Title, gVariablesRegKey, "rcNormalPosition.left", lpwndpl.rcNormalPosition.Left
        SaveSetting App.Title, gVariablesRegKey, "rcNormalPosition.top", lpwndpl.rcNormalPosition.Top
        SaveSetting App.Title, gVariablesRegKey, "rcNormalPosition.right", lpwndpl.rcNormalPosition.Right
        SaveSetting App.Title, gVariablesRegKey, "rcNormalPosition.bottom", lpwndpl.rcNormalPosition.Bottom
    End If

    frmMain.IsVariablesWindowLoaded = False
    frmMain.SetFocus 'VB5 Bug
End Sub

Private Sub Form_Resize()
    Dim CmdLeft As Long
    CmdLeft = Me.ScaleWidth - 1215

    If CmdLeft > 1500 Then
        lstVariables.Width = CmdLeft - 240

        cmdAdd.Left = CmdLeft
        cmdProperties.Left = CmdLeft
        cmdDelete.Left = CmdLeft
    End If
    If Me.ScaleHeight > 1575 Then
        lstVariables.Height = Me.ScaleHeight - 240
    End If

    DoEvents
End Sub

Private Sub lstVariables_DblClick()
    cmdProperties_Click
End Sub

Private Sub cmdAdd_Click()
    Dim NewVariable As New clsVariable
    Set frmVariableProperties.CurrentVariable = NewVariable
    Set frmVariableProperties.VariablesCollection = frmMain.ActiveForm.Variables
    frmVariableProperties.AddNewVariable = True
    PropsChanged = False
    frmVariableProperties.Show vbModal, frmMain

    Screen.MousePointer = vbHourglass
    If PropsChanged Then
        frmMain.ActiveForm.Variables.VariablesCollection.Add NewVariable
        lstVariables.AddItem NewVariable.DisplayName
        lstVariables.ListIndex = lstVariables.ListCount - 1

        'notify the active form of the occured change
        frmMain.ActiveForm.VariableChanged lstVariables.ListCount

        frmMain.ActiveForm.FileChanged = True
        UpdateCommands
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdProperties_Click()
    Dim CurrentVariable As Long
    CurrentVariable = lstVariables.ListIndex + 1

    Set frmVariableProperties.CurrentVariable = frmMain.ActiveForm.Variables.VariablesCollection(CurrentVariable)
    Set frmVariableProperties.VariablesCollection = frmMain.ActiveForm.Variables
    frmVariableProperties.AddNewVariable = False
    PropsChanged = False
    frmVariableProperties.Show vbModal, frmMain

    Screen.MousePointer = vbHourglass
    If PropsChanged Then
        lstVariables.List(lstVariables.ListIndex) = frmMain.ActiveForm.Variables.VariablesCollection(CurrentVariable).DisplayName

        'notify the active form of the occured change
        frmMain.ActiveForm.VariableChanged CurrentVariable

        frmMain.ActiveForm.FileChanged = True
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDelete_Click()
    Dim LastItem As Long
    LastItem = lstVariables.ListIndex

    'is the selected variable in use?
    If frmMain.ActiveForm.IsVariableInUse(frmMain.ActiveForm.Variables.VariablesCollection(LastItem + 1).Name) Then
        MsgBox "Can't delete variable. Variable is in use.", vbExclamation
        Exit Sub
    End If

    'confirmation
    If gblnConfirmDelete Then
        If MsgBox("Are you sure you want to delete this variable?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If


    Screen.MousePointer = vbHourglass

    frmMain.ActiveForm.Variables.VariablesCollection.Remove LastItem + 1
    lstVariables.RemoveItem LastItem

    If LastItem < lstVariables.ListCount Then
        lstVariables.ListIndex = LastItem
    Else
        lstVariables.ListIndex = LastItem - 1
    End If

    frmMain.ActiveForm.FileChanged = True
    UpdateCommands

    Screen.MousePointer = vbDefault
End Sub

Private Sub UpdateCommands()
    If frmMain.ActiveForm Is Nothing Then
        cmdAdd.Enabled = False
        cmdProperties.Enabled = False
        cmdDelete.Enabled = False
    Else
        cmdAdd.Enabled = True 'Unlimited number of variables
        cmdProperties.Enabled = (lstVariables.ListCount > 0)
        cmdDelete.Enabled = cmdProperties.Enabled
    End If
End Sub

Public Sub UpdateList()
    If frmMain.ActiveForm Is Nothing Then
        Me.Caption = "Variables"
        lstVariables.Clear
        UpdateCommands
        Exit Sub
    End If

    Me.Caption = "Variables - " & frmMain.ActiveForm.Caption

    Dim I As Long
    lstVariables.Clear
    For I = 1 To frmMain.ActiveForm.Variables.VariablesCollection.Count
        lstVariables.AddItem frmMain.ActiveForm.Variables.VariablesCollection(I).DisplayName
    Next I
    If lstVariables.ListCount > 0 Then lstVariables.ListIndex = 0

    UpdateCommands
End Sub
