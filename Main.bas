Attribute VB_Name = "modMain"

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Public Const SRCCOPY = &HCC0020
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Const CF_BITMAP = 2
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long

Public Const SW_SHOWNORMAL = 1, SW_SHOWMINIMIZED = 2, SW_SHOWMAXIMIZED = 3, SW_SHOW = 5
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const MAX_PATH = 260
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As Long) As Long

Public Const SHARD_PATH = 2
Public Declare Sub SHAddToRecentDocs Lib "shell32" (ByVal uFlags As Long, ByVal pv As String)

Public Type POINT
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type WINDOWPLACEMENT
    length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINT
    ptMaxPosition As POINT
    rcNormalPosition As RECT
End Type

Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long

Public Const HH_DISPLAY_TOC = 1
Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long

Public Declare Function InitCommonControls Lib "comctl32" () As Long

Global Const EVAL_ERROR = 1, STR_ERROR = "Error" 'MAX_GRAPHS = 32

Global Const gOptionsRegKey As String = "Options", gWindowRegKey As String = "Window", gRecentFunctionsRegKey As String = "RecentFunctions", gRecentVariablesRegKey As String = "RecentVariables", gRecentFilesRegKey As String = "RecentFiles"
Global Const gIntegrateRegKey As String = "Integrate", gEvalRegKey As String = "Evaluate", gGraphPropertiesRegKey As String = "GraphProperties", gGraphsRegKey As String = "Graphs", gVariablePropertiesRegKey As String = "VariableProperties", gVariablesRegKey As String = "Variables", gTraceRegKey As String = "Trace", gAnalyzeRegKey As String = "Analyze"

Global Const gRecentFunctionsCount As Long = 25, gRecentVariablesCount As Long = 10, gGraphMin As Currency = -10, gGraphMax As Currency = 10, gGraphPrecision As Currency = 0.05, gGraphMaxGap As Long = 75, gGraphSteps As Long = 200
Global Const gZoomStep As Long = 250, gDefaultScale As Long = 600, gMoveStep As Long = 1800, gMinScale As Long = 350, gMaxScale As Long = 6600

Global Const gBackgroundColor As Long = vbWhite
Global Const gAxesColor As Long = vbBlack, gAxesStyle As Long = vbSolid
Global Const gGridColor As Long = &HDFDFDF, gGridStyle As Long = vbDot
Global Const gValuesColor As Long = vbBlack, gValuesFontName As String = "MS Sans Serif", gValuesFontSize As Long = 8
Global Const gGraphColor As Long = vbBlue

Global gblnAutoMaximize As Boolean, gblnSaveWindowsPos As Boolean, gblnConfirmDelete As Boolean

Global glngNewWindowNum As Long, gstrRecentFiles(1 To 4) As String

Public Function FindAndActivate(FileName As String) As Boolean
    'Prevent opening the same file twice

    Dim I As Integer, LCFileName As String

    LCFileName = LCase(FileName)
    For I = 1 To Forms.Count - 1
        If Forms(I).MDIChild Then 'is this a MDI child?
            If LCase(Forms(I).FileName) = LCFileName Then 'is this the same file?
                If Forms(I).WindowState = vbMinimized Then
                    If gblnAutoMaximize Then
                        Forms(I).WindowState = vbMaximized
                    Else
                        Forms(I).WindowState = vbNormal
                    End If
                End If
                Forms(I).SetFocus
                FindAndActivate = True
            End If
        End If
    Next I
End Function

Public Function QualifyPath(Path As String) As String
    If Right(Path, 1) <> "\" Then
        QualifyPath = Path + "\"
    Else
        QualifyPath = Path
    End If
End Function

Public Function ZTrim(Buffer As String) As String
    Dim zPos As Integer

    zPos = InStr(Buffer, Chr(0))
    If zPos = 0 Then
        ZTrim = Buffer
    Else
        ZTrim = Left(Buffer, zPos - 1)
    End If
End Function

Public Function GetFileExt(FileName As String) As String
    Dim I As Long, L As Long

    L = Len(FileName)
    For I = L To 1 Step -1
        If Mid(FileName, I, 1) = "." Then
            GetFileExt = Right(FileName, L - I)
            Exit Function
        End If
    Next I
End Function

Public Function GetFileTitle(FileName As String) As String
    Dim I As Long, L As Long

    L = Len(FileName)
    For I = L To 1 Step -1
        If Mid(FileName, I, 1) = "\" Then
            GetFileTitle = Right(FileName, L - I)
            Exit Function
        End If
    Next I
    GetFileTitle = FileName
End Function

Public Function GetFilePath(FileName As String) As String
    Dim I As Long, L As Long

    L = Len(FileName)
    For I = L To 1 Step -1
        If Mid(FileName, I, 1) = "\" Then
            GetFilePath = Left(FileName, I)
            Exit Function
        End If
    Next I
End Function

Public Function ShortenTitle(FileTitle As String) As String
    Dim L As Long, E As Long

    'removing ext
    L = Len(FileTitle)
    E = Len(GetFileExt(FileTitle))
    If E > 0 Then
        FileTitle = Left(FileTitle, L - E - 1)
        L = Len(FileTitle) 'adjusting the changed length
    End If

    'shrinking
    If L > 30 Then
        ShortenTitle = Left(FileTitle, 15) & "..." & Mid(FileTitle, L - 14, 15)
    Else
        ShortenTitle = FileTitle
    End If
End Function

Public Function ErrAssist(Style As Integer) As Integer
    '"Run-time error " & Err.Number
    ErrAssist = MsgBox(Err.Description, Style + vbCritical, Err.Source)
End Function
