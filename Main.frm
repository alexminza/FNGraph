VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "FNGraph"
   ClientHeight    =   4305
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7380
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   120
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "fng"
      Filter          =   "FNGraph Document (*.fng)|*.fng"
      PrinterDefault  =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile_Open 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile_Close 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFile_D1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile_SaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFile_D2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_MRUD 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_MRU 
         Caption         =   "&1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_MRU 
         Caption         =   "&2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_MRU 
         Caption         =   "&3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_MRU 
         Caption         =   "&4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_D3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit_Copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuView_DocumentProperties 
         Caption         =   "&Properties..."
      End
      Begin VB.Menu mnuView_D1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_Variables 
         Caption         =   "&Variables"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuView_Graphs 
         Caption         =   "&Graphs"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTools_Eval 
         Caption         =   "&Evaluate..."
      End
      Begin VB.Menu mnuTools_Trace 
         Caption         =   "&Trace..."
      End
      Begin VB.Menu mnuTools_Analyze 
         Caption         =   "&Analyze..."
      End
      Begin VB.Menu mnuTools_Integrate 
         Caption         =   "&Integrate..."
      End
      Begin VB.Menu mnuTools_D1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTools_Options 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindow_TileH 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuWindow_TileV 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuWindow_Cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindow_ArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp_Topics 
         Caption         =   "Help &Topics"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp_D1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Public IsGraphsWindowLoaded As Boolean, IsVariablesWindowLoaded As Boolean

Public Sub MDIForm_Load()
    Screen.MousePointer = vbArrowHourglass

    'Initialization
    CreateMutex 0, 0, "FNGraphAppMutex"
    InitCommonControls
    App.HelpFile = QualifyPath(App.Path) & "FNGraph.chm"
    cdlMain.flags = cdlPDUseDevModeCopies + cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNFileMustExist

    'Get and apply settings
    GetGlobalSettings 'GetWindowPlacement will show the main window

    'Tool windows sequence
    If gblnSaveWindowsPos Then
        'Graphs
        If GetSetting(App.Title, gGraphsRegKey, "Show", True) Then
            frmGraphs.Show vbModeless, Me
        End If
        'Variables
        If GetSetting(App.Title, gVariablesRegKey, "Show", False) Then
            frmVariables.Show vbModeless, Me
        End If
    End If

    DoEvents

    'Document sequence
    Dim CmdLine As String
    CmdLine = Command

    If CmdLine = "" Then
        FileNew
    Else
        'Stripping quotes
        If (Left(CmdLine, 1) = """") And (Right(CmdLine, 1) = """") Then
            CmdLine = Mid(CmdLine, 2, Len(CmdLine) - 2)
        End If

        Dim strFullPath As String * MAX_PATH
        GetFullPathName CmdLine, MAX_PATH, strFullPath, 0
        CmdLine = ZTrim(strFullPath)

        'Opening file specified in command line
        FileOpen CmdLine
    End If

    DoEvents
    Screen.MousePointer = vbDefault

    'Welcome sequence
    If GetSetting(App.Title, gOptionsRegKey, "ShowWelcome", True) Then
        mnuHelp_About_Click
        SaveSetting App.Title, gOptionsRegKey, "ShowWelcome", False
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveGlobalSettings

    'Save active windows list
    If gblnSaveWindowsPos Then
        SaveSetting App.Title, gGraphsRegKey, "Show", IsGraphsWindowLoaded
        SaveSetting App.Title, gVariablesRegKey, "Show", IsVariablesWindowLoaded
    End If
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Long, FileName As String

    For I = 1 To Data.Files.Count
        FileName = Data.Files(I)
        'Do not open the same file twice
        If Not FindAndActivate(FileName) Then
            FileOpen FileName
        End If
    Next I
End Sub

Private Sub MDIForm_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub mnuFile_New_Click()
    Screen.MousePointer = vbHourglass
    FileNew
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFile_Open_Click()
    On Error GoTo ErrHandler
    Dim FileName As String

    cdlMain.FileName = ""
    cdlMain.ShowOpen
    cdlMain.InitDir = GetFilePath(cdlMain.FileName)
    FileName = cdlMain.FileName
    If Not FindAndActivate(FileName) Then
        FileOpen FileName
    End If
    Exit Sub

ErrHandler:
    If Err.Number = cdlCancel Then Exit Sub
    ErrAssist vbOKOnly
End Sub

Private Sub mnuFile_Close_Click()
    Unload Me.ActiveForm
End Sub

Private Sub mnuFile_Save_Click()
    If Not Me.ActiveForm.FileSaved Then
        FileSaveAs
    Else
        FileSave
    End If
End Sub

Private Sub mnuFile_SaveAs_Click()
    FileSaveAs
End Sub

Private Sub mnuFile_Print_Click()
    On Error GoTo ErrHandler

    cdlMain.ShowPrinter
    'DoEvents

    Screen.MousePointer = vbHourglass
    Printer.Orientation = cdlMain.Orientation
    Printer.Copies = cdlMain.Copies
    Me.ActiveForm.PrintForm

ProcExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHandler:
    If Err.Number = cdlCancel Then Exit Sub
    If ErrAssist(vbRetryCancel) = vbRetry Then Resume
    GoTo ProcExit
End Sub

Private Sub mnuFile_MRU_Click(Index As Integer)
    If Not FindAndActivate(gstrRecentFiles(Index)) Then
        Dim FileName As String

        'create a copy of the file name
        FileName = gstrRecentFiles(Index)
        FileOpen FileName
    End If
End Sub

Private Sub mnuFile_Exit_Click()
    Unload Me
End Sub

Private Sub mnuEdit_Copy_Click()
    Screen.MousePointer = vbHourglass

    Dim ImageX As Long, ImageY As Long
    Dim hChildDC As Long, hMemDC As Long, hBitmap As Long

    hChildDC = Me.ActiveForm.hdc
    ImageX = Me.ActiveForm.ScaleWidth \ Screen.TwipsPerPixelX
    ImageY = Me.ActiveForm.ScaleHeight \ Screen.TwipsPerPixelY

    'make a screenshot
    hMemDC = CreateCompatibleDC(hChildDC)
    hBitmap = CreateCompatibleBitmap(hChildDC, ImageX, ImageY)
    SelectObject hMemDC, hBitmap
    BitBlt hMemDC, 0, 0, ImageX, ImageY, hChildDC, 0, 0, SRCCOPY
    DeleteDC hMemDC

    'set clipboard data
    OpenClipboard Me.hWnd
    EmptyClipboard
    SetClipboardData CF_BITMAP, hBitmap
    CloseClipboard

    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuView_DocumentProperties_Click()
    frmDocumentProperties.Show vbModal, Me
End Sub

Private Sub mnuView_Variables_Click()
    frmVariables.Show vbModeless, Me
End Sub

Private Sub mnuView_Graphs_Click()
    frmGraphs.Show vbModeless, Me
End Sub

Private Sub mnuTools_Eval_Click()
    frmEval.Show vbModal, Me
End Sub

Private Sub mnuTools_Trace_Click()
    frmTrace.Show vbModal, Me
End Sub

Private Sub mnuTools_Analyze_Click()
    frmAnalyze.Show vbModal, Me
End Sub

Private Sub mnuTools_Integrate_Click()
    frmIntegrate.Show vbModal, Me
End Sub

Private Sub mnuTools_Options_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuWindow_TileH_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindow_TileV_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindow_Cascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindow_ArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuHelp_Topics_Click()
    If HtmlHelp(Me.hWnd, App.HelpFile, HH_DISPLAY_TOC, 0) = 0 Then
        MsgBox "Error displaying help. HtmlHelp function failed.", vbCritical
    End If
End Sub

Private Sub mnuHelp_About_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub FileNew()
    Dim frmNewGraph As New frmDocument

    Load frmNewGraph
    frmNewGraph.FileNew

    UpdateFileCmds
End Sub

Private Sub FileOpen(FileName As String)
    On Error GoTo ErrHandler

    If Dir(FileName, vbHidden) = "" Then
        MsgBox "File not found: " & FileName, vbCritical
        Exit Sub
    End If
    If LCase(GetFileExt(FileName)) <> "fng" Then
        MsgBox "Unknown file format: " & FileName, vbCritical
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    Open FileName For Binary Access Read Lock Write As #1

    'Close default document window if it was not modified
    If glngNewWindowNum = 1 Then
        If Not Me.ActiveForm Is Nothing Then
            If Not frmMain.ActiveForm.FileSaved Then
                If Not frmMain.ActiveForm.FileChanged Then
                    Unload Me.ActiveForm
                End If
            End If
        End If
    End If

    Dim frmNewGraph As New frmDocument
    Load frmNewGraph
    frmNewGraph.FileOpen FileName

    AddToFileMRUList FileName
    AddToRecentDocs FileName

    UpdateFileCmds

ProcExit:
    Close #1
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHandler:
    If ErrAssist(vbRetryCancel) = vbRetry Then Resume
    GoTo ProcExit
End Sub

Public Sub SaveActiveFile()
    mnuFile_Save_Click
End Sub

Private Sub FileSave()
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass

    Open Me.ActiveForm.FileName For Output Access Write Lock Write As #1
    Me.ActiveForm.FileSave

ProcExit:
    Close #1
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHandler:
    If ErrAssist(vbRetryCancel) = vbRetry Then Resume
    GoTo ProcExit
End Sub

Private Sub FileSaveAs()
    On Error GoTo ErrHandler
    Dim FileName As String

    cdlMain.FileName = GetFileTitle(Me.ActiveForm.FileName)
    cdlMain.ShowSave
    cdlMain.InitDir = GetFilePath(cdlMain.FileName)
    FileName = cdlMain.FileName
    Me.ActiveForm.FileName = FileName
    FileSave

    Me.ActiveForm.Caption = ShortenTitle(GetFileTitle(FileName))
    AddToFileMRUList FileName
    AddToRecentDocs FileName
    Exit Sub

ErrHandler:
    If Err.Number = cdlCancel Then Exit Sub
    ErrAssist vbOKOnly
End Sub

Private Sub AddToRecentDocs(FileName As String)
    SHAddToRecentDocs SHARD_PATH, FileName
End Sub

Private Sub AddToFileMRUList(FileName As String)
    Dim I As Long, T As Long, LCFileName As String

    LCFileName = LCase(FileName)
    T = 3
    For I = 1 To 4
        If LCase(gstrRecentFiles(I)) = LCFileName Then
            If I = 1 Then Exit Sub
            T = I - 1
            Exit For
        End If
    Next I

    For I = T To 1 Step -1
        If mnuFile_MRU(I).Visible Then
            gstrRecentFiles(I + 1) = gstrRecentFiles(I)
            mnuFile_MRU(I + 1).Caption = "&" & LngToStr(I + 1) & " " & ShortenTitle(GetFileTitle(gstrRecentFiles(I)))
            mnuFile_MRU(I + 1).Visible = True
        End If
    Next I
    gstrRecentFiles(1) = FileName
    mnuFile_MRU(1).Caption = "&1 " & ShortenTitle(GetFileTitle(FileName))

    If Not mnuFile_MRUD.Visible Then
        mnuFile_MRUD.Visible = True
        mnuFile_MRU(1).Visible = True
    End If
End Sub

Private Sub GetFileMRUList()
    Dim I As Long, T As Long, RecentFile As String

    For I = 1 To 4
        RecentFile = GetSetting(App.Title, gRecentFilesRegKey, LngToStr(I), "")
        If RecentFile <> "" Then
            T = T + 1
            gstrRecentFiles(T) = RecentFile
            mnuFile_MRU(T).Caption = "&" & LngToStr(T) & " " & ShortenTitle(GetFileTitle(RecentFile))
            mnuFile_MRU(T).Visible = True
        End If
    Next I
    If T > 0 Then mnuFile_MRUD.Visible = True
End Sub

Private Sub SaveFileMRUList()
    Dim I As Long

    For I = 1 To 4
        SaveSetting App.Title, gRecentFilesRegKey, LngToStr(I), gstrRecentFiles(I)
    Next I
End Sub

Public Sub UpdateGraphsList()
    If IsGraphsWindowLoaded Then
        frmGraphs.UpdateList
    End If
    If IsVariablesWindowLoaded Then
        frmVariables.UpdateList
    End If
End Sub

Public Sub UpdateGraphCmds()
    Dim TUpdate As Boolean

    If frmMain.ActiveForm Is Nothing Then
        TUpdate = False
    Else
        TUpdate = (Me.ActiveForm.Graphs.Count > 0)
    End If

    mnuTools_Trace.Enabled = TUpdate
    mnuTools_Analyze.Enabled = TUpdate
    mnuTools_Integrate.Enabled = TUpdate
End Sub

Public Sub UpdateFileCmds()
    Dim TUpdate As Boolean
    TUpdate = Not (frmMain.ActiveForm Is Nothing)

    mnuFile_Save.Enabled = TUpdate
    mnuFile_SaveAs.Enabled = TUpdate
    mnuFile_Close.Enabled = TUpdate
    mnuFile_Print.Enabled = TUpdate
    mnuEdit_Copy.Enabled = TUpdate
    mnuView_DocumentProperties.Enabled = TUpdate
    mnuWindow_ArrangeIcons.Enabled = TUpdate
    mnuWindow_Cascade.Enabled = TUpdate
    mnuWindow_TileH.Enabled = TUpdate
    mnuWindow_TileV.Enabled = TUpdate

    UpdateGraphCmds
End Sub

Private Sub SaveGlobalSettings()
    Dim lpwndpl As WINDOWPLACEMENT

    lpwndpl.length = Len(lpwndpl)
    GetWindowPlacement Me.hWnd, lpwndpl
    SaveSetting App.Title, gWindowRegKey, "flags", lpwndpl.flags
    SaveSetting App.Title, gWindowRegKey, "showCmd", lpwndpl.showCmd
    SaveSetting App.Title, gWindowRegKey, "ptMaxPosition.X", lpwndpl.ptMaxPosition.X
    SaveSetting App.Title, gWindowRegKey, "ptMaxPosition.Y", lpwndpl.ptMaxPosition.Y
    SaveSetting App.Title, gWindowRegKey, "ptMinPosition.X", lpwndpl.ptMinPosition.X
    SaveSetting App.Title, gWindowRegKey, "ptMinPosition.Y", lpwndpl.ptMinPosition.Y
    SaveSetting App.Title, gWindowRegKey, "rcNormalPosition.left", lpwndpl.rcNormalPosition.Left
    SaveSetting App.Title, gWindowRegKey, "rcNormalPosition.top", lpwndpl.rcNormalPosition.Top
    SaveSetting App.Title, gWindowRegKey, "rcNormalPosition.right", lpwndpl.rcNormalPosition.Right
    SaveSetting App.Title, gWindowRegKey, "rcNormalPosition.bottom", lpwndpl.rcNormalPosition.Bottom

    'SaveSetting App.Title, gOptionsRegKey, "ViewToolbar", mnuView_Toolbar.Checked
    'SaveSetting App.Title, gOptionsRegKey, "ViewStatusBar", mnuView_StatusBar.Checked
    SaveSetting App.Title, gOptionsRegKey, "AutoMaximize", gblnAutoMaximize
    SaveSetting App.Title, gOptionsRegKey, "SaveWindowsPos", gblnSaveWindowsPos
    SaveSetting App.Title, gOptionsRegKey, "ConfirmDelete", gblnConfirmDelete
    SaveSetting App.Title, gOptionsRegKey, "RecentFolder", cdlMain.InitDir

    SaveFileMRUList
End Sub

Private Sub GetGlobalSettings()
    Dim lpwndpl As WINDOWPLACEMENT

    lpwndpl.length = Len(lpwndpl)
    lpwndpl.flags = GetSetting(App.Title, gWindowRegKey, "flags", 0)
    lpwndpl.showCmd = GetSetting(App.Title, gWindowRegKey, "showCmd", SW_SHOWMAXIMIZED)
    If lpwndpl.showCmd = SW_SHOWMINIMIZED Then lpwndpl.showCmd = SW_SHOWMAXIMIZED
    lpwndpl.ptMaxPosition.X = GetSetting(App.Title, gWindowRegKey, "ptMaxPosition.X", -1)
    lpwndpl.ptMaxPosition.Y = GetSetting(App.Title, gWindowRegKey, "ptMaxPosition.Y", -1)
    lpwndpl.ptMinPosition.X = GetSetting(App.Title, gWindowRegKey, "ptMinPosition.X", -1)
    lpwndpl.ptMinPosition.Y = GetSetting(App.Title, gWindowRegKey, "ptMinPosition.Y", -1)
    lpwndpl.rcNormalPosition.Left = GetSetting(App.Title, gWindowRegKey, "rcNormalPosition.left", 0)
    lpwndpl.rcNormalPosition.Top = GetSetting(App.Title, gWindowRegKey, "rcNormalPosition.top", 0)
    lpwndpl.rcNormalPosition.Right = GetSetting(App.Title, gWindowRegKey, "rcNormalPosition.right", 640)
    lpwndpl.rcNormalPosition.Bottom = GetSetting(App.Title, gWindowRegKey, "rcNormalPosition.bottom", 480)
    SetWindowPlacement Me.hWnd, lpwndpl 'this shows the main window

    'mnuView_Toolbar.Checked = GetSetting(App.Title, gOptionsRegKey, "ViewToolbar", True)
    'mnuView_StatusBar.Checked = GetSetting(App.Title, gOptionsRegKey, "ViewStatusBar", True)
    gblnAutoMaximize = GetSetting(App.Title, gOptionsRegKey, "AutoMaximize", True)
    gblnSaveWindowsPos = GetSetting(App.Title, gOptionsRegKey, "SaveWindowsPos", True)
    gblnConfirmDelete = GetSetting(App.Title, gOptionsRegKey, "ConfirmDelete", True)
    cdlMain.InitDir = GetSetting(App.Title, gOptionsRegKey, "RecentFolder")

    GetFileMRUList
End Sub
