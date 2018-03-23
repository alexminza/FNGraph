VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "MS Shell Dlg"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraPreferences 
      Caption         =   "Preferences"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CheckBox chkConfirmDelete 
         Caption         =   "Confirm delete"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox chkSaveDialogsPos 
         Caption         =   "Save &dialogs positions"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox chkAutoMaximize 
         Caption         =   "&Auto maximize windows"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Defaults"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
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
    chkAutoMaximize.Value = IIf(gblnAutoMaximize, vbChecked, vbUnchecked)
    chkSaveDialogsPos.Value = IIf(gblnSaveWindowsPos, vbChecked, vbUnchecked)
    chkConfirmDelete.Value = IIf(gblnConfirmDelete, vbChecked, vbUnchecked)
End Sub

Private Sub cmdOK_Click()
    gblnAutoMaximize = (chkAutoMaximize.Value = vbChecked)
    gblnSaveWindowsPos = (chkSaveDialogsPos.Value = vbChecked)
    gblnConfirmDelete = (chkConfirmDelete.Value = vbChecked)

    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDefaults_Click()
    chkAutoMaximize.Value = vbChecked
    chkSaveDialogsPos.Value = vbChecked
    chkConfirmDelete.Value = vbChecked
End Sub
