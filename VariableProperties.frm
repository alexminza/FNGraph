VERSION 5.00
Begin VB.Form frmVariableProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Variable Properties"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2535
   BeginProperty Font 
      Name            =   "MS Shell Dlg"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VariableProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboVariableName 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtVariableValue 
      Height          =   285
      Left            =   840
      MaxLength       =   12
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblVariableValue 
      Caption         =   "&Value:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblVariableName 
      Caption         =   "&Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmVariableProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public CurrentVariable As clsVariable, VariablesCollection As clsVariablesCollection, AddNewVariable As Boolean

Private Sub Form_Load()
    'Set window position
    If gblnSaveWindowsPos Then
        Me.Top = GetSetting(App.Title, gVariablePropertiesRegKey, "Top", (Screen.Height - Me.Height) \ 2)
        Me.Left = GetSetting(App.Title, gVariablePropertiesRegKey, "Left", (Screen.Width - Me.Width) \ 2)
    Else
        Me.Top = (Screen.Height - Me.Height) \ 2
        Me.Left = (Screen.Width - Me.Width) \ 2
    End If

    'Load recent variable names
    Dim I As Long, S As String
    For I = 1 To gRecentVariablesCount
        S = GetSetting(App.Title, gRecentVariablesRegKey, LngToStr(I), "")
        If S <> "" Then cboVariableName.AddItem S
    Next I

    'Add or Properties
    If AddNewVariable Then
        Me.Caption = "Add Variable"
    Else
        Me.Caption = "Variable Properties"
    End If

    cboVariableName.Text = CurrentVariable.Name
    txtVariableValue.Text = CurToStr(CurrentVariable.Value)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Save window position
    If gblnSaveWindowsPos Then
        SaveSetting App.Title, gVariablePropertiesRegKey, "Top", Me.Top
        SaveSetting App.Title, gVariablePropertiesRegKey, "Left", Me.Left
    End If
End Sub

Private Sub Form_Activate()
    'if changing properties - default to variable value field
    If Not AddNewVariable Then
        txtVariableValue.SetFocus
    End If
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandler
    Dim blnPropsChanged As Boolean

    'Validate Name
    cboVariableName.Text = PrepareExpression(cboVariableName.Text)
    If cboVariableName.Text = "" Then
        MsgBox "Invalid variable name.", vbExclamation
        cboVariableName.SetFocus
        Exit Sub
    End If
    If CheckVariableName(cboVariableName.Text) Then
        cboVariableName.SetFocus
        Exit Sub
    End If
    If CurrentVariable.Name <> cboVariableName.Text Then
        If frmMain.ActiveForm.IsVariableInUse(CurrentVariable.Name) Then
            MsgBox "Can't rename variable. Variable is in use.", vbExclamation
            cboVariableName.Text = CurrentVariable.Name
            'set focus to the variable value text box
            txtVariableValue_GotFocus
            txtVariableValue.SetFocus
            Exit Sub
        End If
        If VariablesCollection.IsDefined(cboVariableName.Text) Then
            MsgBox "Variable already defined.", vbExclamation
            txtVariableValue_GotFocus
            cboVariableName.SetFocus
            Exit Sub
        End If

        'set the variable name property
        CurrentVariable.Name = cboVariableName.Text
        SaveRecentVariables
        blnPropsChanged = True
    End If

    'Validate Value
    Dim curTemp As Currency
    curTemp = StrToCur(txtVariableValue.Text)
    If CurrentVariable.Value <> curTemp Then
        CurrentVariable.Value = curTemp
        blnPropsChanged = True
    End If

    frmVariables.PropsChanged = blnPropsChanged
    Unload Me
    Exit Sub

ErrHandler:
    ErrAssist vbOKOnly
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtVariableValue_GotFocus()
    txtVariableValue.SelStart = 0
    txtVariableValue.SelLength = Len(txtVariableValue.Text)
End Sub

Private Sub SaveRecentVariables()
    Dim I As Long, T As Long

    'save current variable
    SaveSetting App.Title, gRecentVariablesRegKey, "1", cboVariableName.Text

    T = 1
    For I = 0 To cboVariableName.ListCount - 1
        If cboVariableName.Text <> cboVariableName.List(I) Then
            T = T + 1
            SaveSetting App.Title, gRecentVariablesRegKey, LngToStr(T), cboVariableName.List(I)

            If T = gRecentVariablesCount Then Exit For
        End If
    Next I
End Sub

