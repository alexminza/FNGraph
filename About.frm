VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About FNGraph"
   ClientHeight    =   2895
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4455
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Shell Dlg"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1998.18
   ScaleMode       =   0  'User
   ScaleWidth      =   4183.475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   1140
   End
   Begin VB.Label lblDevSnapshot 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Development Snapshot"
      Height          =   195
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   1665
   End
   Begin VB.Label lblAuthor 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Copyright © 2001-2002 Alexander Minza"
      Height          =   195
      Left            =   1260
      TabIndex        =   1
      Top             =   1800
      Width           =   2955
   End
   Begin VB.Label lblURL 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "www.ournet.md/~fngraph"
      BeginProperty Font 
         Name            =   "MS Shell Dlg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2295
      MouseIcon       =   "About.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1440
      Width           =   1920
   End
   Begin VB.Label lblEmail 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "alex_minza@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Shell Dlg"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2415
      MouseIcon       =   "About.frx":0162
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   240
      Picture         =   "About.frx":02B8
      Top             =   240
      Width           =   480
   End
   Begin VB.Line linDivider 
      BorderColor     =   &H80000010&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   109.869
      X2              =   4060.459
      Y1              =   1521.93
      Y2              =   1521.93
   End
   Begin VB.Label lblTitle 
      Caption         =   "FNGraph"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3525
   End
   Begin VB.Line linDivider 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   109.869
      X2              =   4060.459
      Y1              =   1532.283
      Y2              =   1532.283
   End
End
Attribute VB_Name = "frmAbout"
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
    lblTitle.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub lblEmail_Click()
    Screen.MousePointer = vbArrowHourglass
    ShellExecute Me.hWnd, "open", "mailto:Alexander%20Minza%3calex_minza@hotmail.com%3e?subject=" & lblTitle.Caption, vbNullString, vbNullString, SW_SHOWNORMAL
    Screen.MousePointer = vbDefault
End Sub

Private Sub lblEmail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEmail.ForeColor = vbRed
End Sub

Private Sub lblEmail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEmail.ForeColor = vbBlue
    DoEvents
End Sub

Private Sub lblURL_Click()
    Screen.MousePointer = vbArrowHourglass
    ShellExecute Me.hWnd, "open", "http://www.ournet.md/~fngraph", vbNullString, vbNullString, SW_SHOWNORMAL
    Screen.MousePointer = vbDefault
End Sub

Private Sub lblURL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblURL.ForeColor = vbRed
End Sub

Private Sub lblURL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblURL.ForeColor = vbBlue
    DoEvents
End Sub
