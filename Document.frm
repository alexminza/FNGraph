VERSION 5.00
Begin VB.Form frmDocument 
   AutoRedraw      =   -1  'True
   Caption         =   "New Document"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "Document.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Tag             =   "0"
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Copyright © 2001-2002 Alexander Minza

'alex_minza@hotmail.com
'http://www.ournet.md/~fngraph
'http://www.hi-tech.ournet.md

Option Explicit

Public plngDocumentScaleX As Long, plngDocumentScaleY As Long, plngBackgroundColor As Long
Public pblnAxesVisible As Boolean, plngAxesColor As Long, plngAxesStyle As Long, plngAxesWidth As Long
Public pblnGridVisible As Boolean, plngGridColor As Long, plngGridStyle As Long, plngGridWidth As Long
Public pblnValuesVisible As Boolean, plngValuesColor As Long, pstrValuesFontName As String, plngValuesFontSize As Long, pblnValuesFontBold As Boolean, pblnValuesFontItalic As Boolean

Public FileName As String, FileChanged As Boolean, FileSaved As Boolean, ResizeEnabled As Boolean

Private lngMouseDownX As Long, lngMouseDownY As Long, lngMouseDownScaleX As Long, lngMouseDownScaleY As Long, lngMouseDownOffsetX As Long, lngMouseDownOffsetY As Long
Private llngX0 As Long, llngY0 As Long, llngSW As Long, llngSH As Long, lngMouseOperation As Long, lngTemp As Long

Private Const TagDepthMax = 3
Private TagStack(1 To TagDepthMax) As String, TagDepth As Long

Public Graphs As New Collection 'of clsGraph
Public Variables As New clsVariablesCollection

Public Sub GraphsTrace(GraphIndex As Long, Pos As Long)
    On Error GoTo ErrHandler

    Dim CX As Long, CY As Long, EvX As Currency, EvY As Currency, EvStat As Long
    Graphs(GraphIndex).EvVal Pos, EvX, EvY, EvStat
    CX = llngX0 + EvX * plngDocumentScaleX
    CY = llngY0 - EvY * plngDocumentScaleY

    'Visible area is (0, 0) - (Me.ScaleWidth - 1, Me.ScaleHeight - 1)
    If (CX >= 0) And (CX < llngSW) And (CY >= 0) And (CY < llngSH) Then
        Me.Line (CX - 200, CY)-Step(400, 0), vbBlack
        Me.Line (CX, CY - 200)-Step(0, 400), vbBlack
    End If
    Exit Sub

ErrHandler:
End Sub

Public Sub GraphsAdd()
    Dim NewGraph As New clsGraph

    'set reference to document variables
    Set NewGraph.Expression.Variables = Variables

    'set default values
    NewGraph.Min = gGraphMin
    NewGraph.Max = gGraphMax
    NewGraph.Precision = gGraphPrecision
    NewGraph.MaxGap = gGraphMaxGap
    NewGraph.Color = gGraphColor
    NewGraph.DrawLines = True
    NewGraph.DrawPoints = False
    NewGraph.LinesWidth = 1
    NewGraph.PointsSize = 1
    NewGraph.Visible = True

    Graphs.Add NewGraph
End Sub

Public Sub VariableChanged(VariableIndex As Long)
    Dim I As Long, VariableName As String, HasChanges As Boolean

    VariableName = Variables.VariablesCollection(VariableIndex).Name
    For I = 1 To Graphs.Count
        If Graphs(I).Expression.UsesIdentifier(VariableName) Then
            Graphs(I).CacheValues
            HasChanges = True
        End If
    Next I

    If HasChanges Then
        WindowRefresh
    End If
End Sub

Public Function IsVariableInUse(VariableName As String)
    Dim I As Long

    For I = 1 To Graphs.Count
        If Graphs(I).Expression.UsesIdentifier(VariableName) Then
            IsVariableInUse = True
            Exit Function
        End If
    Next I
End Function

Private Sub SetDefaultDocumentProperties()
    plngDocumentScaleX = gDefaultScale
    plngDocumentScaleY = gDefaultScale
    llngSW = Me.ScaleWidth
    llngSH = Me.ScaleHeight
    llngX0 = llngSW / 2
    llngY0 = llngSH / 2

    plngBackgroundColor = gBackgroundColor
    pblnAxesVisible = True
    plngAxesColor = gAxesColor
    plngAxesStyle = gAxesStyle
    plngAxesWidth = 1
    pblnGridVisible = True
    plngGridColor = gGridColor
    plngGridStyle = gGridStyle
    plngGridWidth = 1
    pblnValuesVisible = True
    plngValuesColor = gValuesColor
    pstrValuesFontName = gValuesFontName
    plngValuesFontSize = gValuesFontSize
    pblnValuesFontBold = False
    pblnValuesFontItalic = False
End Sub

Private Sub Form_Load()
    If gblnAutoMaximize Then
        Me.WindowState = vbMaximized
    End If

    Me.Show
    SetDefaultDocumentProperties
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If FileChanged Then
        Select Case MsgBox("Save chages to " & FileName & "?", vbQuestion + vbYesNoCancel)
            Case vbYes: frmMain.SaveActiveFile
            Case vbCancel: Cancel = True
        End Select
    End If
End Sub

Private Sub Form_Terminate()
    frmMain.UpdateFileCmds
    frmMain.UpdateGraphsList
End Sub

Private Sub Form_Resize()
    If ResizeEnabled Then
        llngX0 = llngX0 + (Me.ScaleWidth - llngSW) / 2
        llngY0 = llngY0 + (Me.ScaleHeight - llngSH) / 2
        llngSW = Me.ScaleWidth
        llngSH = Me.ScaleHeight
        WindowRefresh
    End If
End Sub

Private Sub Form_Activate()
    frmMain.UpdateGraphsList
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Using the form level lngTemp variable for temporary calculations
    Select Case KeyCode
        Case vbKeyLeft, vbKeyNumpad4:
            Select Case Shift
                Case vbCtrlMask 'ZoomOutX
                    If plngDocumentScaleX > gMinScale Then
                        lngTemp = plngDocumentScaleX - gZoomStep
                        If lngTemp < gMinScale Then lngTemp = gMinScale

                        llngX0 = llngX0 + ((llngSW / 2 - llngX0) / plngDocumentScaleX * (plngDocumentScaleX - lngTemp))
                        plngDocumentScaleX = lngTemp
                        WindowRefresh
                    End If
                Case 0 'MoveLeft
                    llngX0 = llngX0 + gMoveStep
                    WindowRefresh
            End Select

        Case vbKeyRight, vbKeyNumpad6:
            Select Case Shift
                Case vbCtrlMask 'ZoomInX
                    If plngDocumentScaleX < gMaxScale Then
                        lngTemp = plngDocumentScaleX + gZoomStep
                        If lngTemp > gMaxScale Then lngTemp = gMaxScale

                        llngX0 = llngX0 - ((llngSW / 2 - llngX0) / plngDocumentScaleX * (lngTemp - plngDocumentScaleX))
                        plngDocumentScaleX = lngTemp
                        WindowRefresh
                    End If
                Case 0 'MoveRight
                    llngX0 = llngX0 - gMoveStep
                    WindowRefresh
            End Select

        Case vbKeyUp, vbKeyNumpad8:
            Select Case Shift
                Case vbCtrlMask 'ZoomInY
                    If plngDocumentScaleY < gMaxScale Then
                        lngTemp = plngDocumentScaleY + gZoomStep
                        If lngTemp > gMaxScale Then lngTemp = gMaxScale

                        llngY0 = llngY0 + ((llngY0 - llngSH / 2) / plngDocumentScaleY * (lngTemp - plngDocumentScaleY))
                        plngDocumentScaleY = lngTemp
                        WindowRefresh
                    End If
                Case 0 'MoveUp
                    llngY0 = llngY0 + gMoveStep
                    WindowRefresh
            End Select

        Case vbKeyDown, vbKeyNumpad2:
            Select Case Shift
                Case vbCtrlMask 'ZoomOutY
                    If plngDocumentScaleY > gMinScale Then
                        lngTemp = plngDocumentScaleY - gZoomStep
                        If lngTemp < gMinScale Then lngTemp = gMinScale

                        llngY0 = llngY0 - ((llngY0 - llngSH / 2) / plngDocumentScaleY * (plngDocumentScaleY - lngTemp))
                        plngDocumentScaleY = lngTemp
                        WindowRefresh
                    End If
                Case 0 'MoveDown
                    llngY0 = llngY0 - gMoveStep
                    WindowRefresh
            End Select

        Case vbKeyInsert, vbKeyNumpad0: 'MoveCenter
            Dim X As Long, Y As Long

            X = llngSW \ 2
            Y = llngSH \ 2

            If (llngX0 <> X) Or (llngY0 <> Y) Then
                llngX0 = X
                llngY0 = Y
                WindowRefresh
            End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Using the form level lngTemp variable for temporary calculations
    Select Case KeyAscii
        Case 10: 'Ctrl + Enter
            If (plngDocumentScaleX <> gDefaultScale) Or (plngDocumentScaleY <> gDefaultScale) Then
                llngX0 = llngX0 - ((llngSW / 2 - llngX0) / plngDocumentScaleX * (gDefaultScale - plngDocumentScaleX))
                plngDocumentScaleX = gDefaultScale

                llngY0 = llngY0 + ((llngY0 - llngSH / 2) / plngDocumentScaleY * (gDefaultScale - plngDocumentScaleY))
                plngDocumentScaleY = gDefaultScale

                WindowRefresh
            End If

        Case 13: 'Enter
            If plngDocumentScaleX <> plngDocumentScaleY Then
                lngTemp = (plngDocumentScaleX + plngDocumentScaleY) / 2

                llngX0 = llngX0 - ((llngSW / 2 - llngX0) / plngDocumentScaleX * (lngTemp - plngDocumentScaleX))
                plngDocumentScaleX = lngTemp

                llngY0 = llngY0 + ((llngY0 - llngSH / 2) / plngDocumentScaleY * (lngTemp - plngDocumentScaleY))
                plngDocumentScaleY = lngTemp

                WindowRefresh
            End If

        Case 43: '+
            If (plngDocumentScaleX < gMaxScale) And (plngDocumentScaleY < gMaxScale) Then
                lngTemp = plngDocumentScaleX + gZoomStep
                If lngTemp > gMaxScale Then lngTemp = gMaxScale
                llngX0 = llngX0 - ((llngSW / 2 - llngX0) / plngDocumentScaleX * (lngTemp - plngDocumentScaleX))
                plngDocumentScaleX = lngTemp

                lngTemp = plngDocumentScaleY + gZoomStep
                If lngTemp > gMaxScale Then lngTemp = gMaxScale
                llngY0 = llngY0 + ((llngY0 - llngSH / 2) / plngDocumentScaleY * (lngTemp - plngDocumentScaleY))
                plngDocumentScaleY = lngTemp

                WindowRefresh
            End If

        Case 45: '-
            If (plngDocumentScaleX > gMinScale) And (plngDocumentScaleY > gMinScale) Then
                lngTemp = plngDocumentScaleX - gZoomStep
                If lngTemp < gMinScale Then lngTemp = gMinScale
                llngX0 = llngX0 + ((llngSW / 2 - llngX0) / plngDocumentScaleX * (plngDocumentScaleX - lngTemp))
                plngDocumentScaleX = lngTemp

                lngTemp = plngDocumentScaleY - gZoomStep
                If lngTemp < gMinScale Then lngTemp = gMinScale
                llngY0 = llngY0 - ((llngY0 - llngSH / 2) / plngDocumentScaleY * (plngDocumentScaleY - lngTemp))
                plngDocumentScaleY = lngTemp

                WindowRefresh
            End If
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Select Case Shift
            Case 0
                lngMouseOperation = 1 'Moving
                lngMouseDownOffsetX = X - llngX0
                lngMouseDownOffsetY = Y - llngY0
                'Me.MousePointer = vbSizeAll

            Case vbCtrlMask
                lngMouseOperation = 2 'Zooming
                lngMouseDownX = X
                lngMouseDownY = Y
                lngMouseDownScaleX = plngDocumentScaleX
                lngMouseDownScaleY = plngDocumentScaleY
        End Select
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Select Case lngMouseOperation
            Case 1 'Moving
                llngX0 = X - lngMouseDownOffsetX
                llngY0 = Y - lngMouseDownOffsetY
                WindowRefresh

            Case 2 'Zooming
                lngTemp = lngMouseDownScaleX + X - lngMouseDownX
                If lngTemp < gMinScale Then
                    lngTemp = gMinScale
                ElseIf lngTemp > gMaxScale Then
                    lngTemp = gMaxScale
                End If

                llngX0 = llngX0 - ((llngSW / 2 - llngX0) / plngDocumentScaleX * (lngTemp - plngDocumentScaleX))
                plngDocumentScaleX = lngTemp

                lngTemp = lngMouseDownScaleY + lngMouseDownY - Y
                If lngTemp < gMinScale Then
                    lngTemp = gMinScale
                ElseIf lngTemp > gMaxScale Then
                    lngTemp = gMaxScale
                End If

                llngY0 = llngY0 + ((llngY0 - llngSH / 2) / plngDocumentScaleY * (lngTemp - plngDocumentScaleY))
                plngDocumentScaleY = lngTemp

                WindowRefresh
        End Select
    'Else
        'Show current coordinates in the status bar
        'frmMain.sbrMain.PanelText(1) = CurToStr(Int(((X - llngX0) / plngDocumentScale) * 10 + 0.5) / 10) & " x " & CurToStr(Int(((llngY0 - Y) / plngDocumentScale) * 10 + 0.5) / 10)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbLeftButton
                lngMouseOperation = 0
                'Me.MousePointer = vbDefault
        Case vbRightButton
            If Not frmMain.ActiveForm Is Me Then
                Me.SetFocus
                DoEvents
            End If
            PopupMenu frmMain.mnuView
    End Select
End Sub

Private Sub DrawAxes()
    'Viewable area is (0, 0) - (Me.ScaleWidth - 1, Me.ScaleHeight - 1)
    Me.DrawStyle = plngAxesStyle
    Me.DrawWidth = plngAxesWidth

    If (llngX0 >= 0) And (llngX0 < llngSW) Then
        Me.Line (llngX0, 0)-Step(0, llngSH), plngAxesColor
    End If

    If (llngY0 >= 0) And (llngY0 < llngSH) Then
        Me.Line (0, llngY0)-Step(llngSW, 0), plngAxesColor
    End If
End Sub

Private Sub DrawValues()
    'needs optimization
    Me.FontName = pstrValuesFontName
    Me.FontSize = plngValuesFontSize
    Me.FontBold = pblnValuesFontBold
    Me.FontItalic = pblnValuesFontItalic

    Dim I As Long, ST As String, IStart As Long
    Dim ClipX0 As Long, ClipY0 As Long, VisibleX As Boolean, VisibleY As Boolean
    Dim ClipLeft As Long, ClipRight As Long, ClipTop As Long, ClipBottom As Long, Pad As Long

    Me.ForeColor = plngValuesColor
    Pad = plngAxesWidth * 7.5

    ClipLeft = -plngDocumentScaleX
    ClipRight = llngSW + plngDocumentScaleX
    ClipTop = -plngDocumentScaleY
    ClipBottom = llngSH + plngDocumentScaleY

    VisibleX = (llngX0 >= ClipLeft) And (llngX0 <= ClipRight)
    VisibleY = (llngY0 >= ClipTop) And (llngY0 <= ClipBottom)

    Select Case llngX0
        Case Is > llngSW
            ClipX0 = llngSW + (llngX0 - llngSW) Mod plngDocumentScaleX
        Case Is < 0
            ClipX0 = llngX0 Mod plngDocumentScaleX
        Case Else
            ClipX0 = llngX0
    End Select

    Select Case llngY0
        Case Is > llngSH
            ClipY0 = llngSH + (llngY0 - llngSH) Mod plngDocumentScaleY
        Case Is < 0
            ClipY0 = llngY0 Mod plngDocumentScaleY
        Case Else
            ClipY0 = llngY0
    End Select

    'X
    If VisibleY Then
        If llngX0 <= ClipRight Then
            IStart = ClipX0 - plngDocumentScaleX
        Else
            IStart = ClipX0
        End If

        '-X
        For I = IStart To ClipLeft Step -plngDocumentScaleX
            ST = "-" & LngToStr((llngX0 - I) \ plngDocumentScaleX)
            Me.CurrentX = I - Me.TextWidth(ST) \ 2
            Me.CurrentY = llngY0 + Pad
            Me.Print ST
        Next I

        If llngX0 >= ClipLeft Then
            IStart = ClipX0 + plngDocumentScaleX
        Else
            IStart = ClipX0
        End If

        '+X
        For I = IStart To ClipRight Step plngDocumentScaleX
            ST = LngToStr((I - llngX0) \ plngDocumentScaleX)
            Me.CurrentX = I - Me.TextWidth(ST) \ 2
            Me.CurrentY = llngY0 + Pad
            Me.Print ST
        Next I
    End If

    'Y
    If VisibleX Then
        If llngY0 <= ClipBottom Then
            IStart = ClipY0 - plngDocumentScaleY
        Else
            IStart = ClipY0
        End If

        '+Y
        For I = IStart To ClipTop Step -plngDocumentScaleY
            ST = LngToStr((llngY0 - I) \ plngDocumentScaleY)
            Me.CurrentX = llngX0 - TextWidth(ST) - Pad - 30
            Me.CurrentY = I - Me.TextHeight(ST) \ 2
            Me.Print ST
        Next I

        If llngY0 >= ClipTop Then
            IStart = ClipY0 + plngDocumentScaleY
        Else
            IStart = ClipY0
        End If

        '-Y
        For I = IStart To ClipBottom Step plngDocumentScaleY
            ST = "-" & LngToStr((I - llngY0) \ plngDocumentScaleY)
            Me.CurrentX = llngX0 - TextWidth(ST) - Pad - 30
            Me.CurrentY = I - Me.TextHeight(ST) \ 2
            Me.Print ST
        Next I
    End If

    '0
    If VisibleX And VisibleY Then
        Me.CurrentX = llngX0 - Me.TextWidth("0") - Pad - 30
        Me.CurrentY = llngY0 + Pad
        Me.Print "0"
    End If
End Sub

Private Sub DrawGrid()
    Dim I As Long, IStart As Long

    Me.DrawStyle = plngGridStyle
    Me.DrawWidth = plngGridWidth

    '-X
    If llngX0 > llngSW Then
        IStart = llngSW + (llngX0 - llngSW) Mod plngDocumentScaleX
    Else
        IStart = llngX0
    End If

    For I = IStart To 0 Step -plngDocumentScaleX
        Me.Line (I, 0)-Step(0, llngSH), plngGridColor
    Next I

    '+X
    If llngX0 < 0 Then
        IStart = llngX0 Mod plngDocumentScaleX
    Else
        IStart = llngX0
    End If

    For I = IStart To llngSW Step plngDocumentScaleX
        Me.Line (I, 0)-Step(0, llngSH), plngGridColor
    Next I

    '+Y
    If llngY0 > llngSH Then
        IStart = llngSH + (llngY0 - llngSH) Mod plngDocumentScaleY
    Else
        IStart = llngY0
    End If

    For I = IStart To 0 Step -plngDocumentScaleY
        Me.Line (0, I)-Step(llngSW, 0), plngGridColor
    Next I

    '-Y
    If llngY0 < 0 Then
        IStart = llngY0 Mod plngDocumentScaleY
    Else
        IStart = llngY0
    End If

    For I = IStart To llngSH Step plngDocumentScaleY
        Me.Line (0, I)-Step(llngSW, 0), plngGridColor
    Next I
End Sub

Private Sub GraphDrawPoints(GraphIndex As Long)
    Dim CX As Long, CY As Long, ClipRight As Long, ClipBottom As Long
    Dim X As Long, C As Long, T As Long, EvX As Currency, EvY As Currency, EvStat As Long

    T = Graphs(GraphIndex).ValuesCount
    C = Graphs(GraphIndex).Color

    Me.DrawWidth = Graphs(GraphIndex).PointsSize

    'Visible area is (0, 0) - (Me.ScaleWidth - 1, Me.ScaleHeight - 1)
    ClipRight = llngSW - 1
    ClipBottom = llngSH - 1

    On Error Resume Next
    For X = 1 To T
        Graphs(GraphIndex).EvVal X, EvX, EvY, EvStat
        If EvStat <> 1 Then 'EVAL_ERROR
            CX = llngX0 + EvX * plngDocumentScaleX
            CY = llngY0 - EvY * plngDocumentScaleY

            If (CX >= 0) And (CX <= ClipRight) And (CY >= 0) And (CY <= ClipBottom) Then
                PSet (CX, CY), C
            End If
        End If
    Next X
End Sub

Private Function CohenSutherlandComputeCode(X As Long, Y As Long, ClipRight As Long, ClipBottom As Long) As Long
    Dim CSCode As Long

    If Y > ClipBottom Then
        CSCode = CSCode Or 1 'Top
    ElseIf Y < 0 Then
        CSCode = CSCode Or 2 'Bottom
    End If

    If X > ClipRight Then
        CSCode = CSCode Or 4 'Right
    ElseIf X < 0 Then
        CSCode = CSCode Or 8 'Left
    End If

    CohenSutherlandComputeCode = CSCode
End Function

Private Sub GraphDrawLines(GraphIndex As Long)
    Dim X As Long, T As Long, M As Long, C As Long, EvX As Currency, EvY As Currency, EvStat As Long
    Dim PrevEvY As Long, PrevEvStat As Long
    Dim LX As Long, LY As Long, CX As Long, CY As Long 'calculated coordinates
    Dim ClipRight As Long, ClipBottom As Long, CSCodeL As Long, CSCodeC As Long, CSCode As Long
    Dim SX As Long, SY As Long 'computed clip coordinates
    Dim SLX As Long, SLY As Long, SCX As Long, SCY As Long 'clipped coordinates

    T = Graphs(GraphIndex).ValuesCount
    M = Graphs(GraphIndex).MaxGap
    C = Graphs(GraphIndex).Color
    PrevEvStat = EVAL_ERROR

    Me.DrawStyle = vbSolid
    Me.DrawWidth = Graphs(GraphIndex).LinesWidth

    ClipRight = llngSW
    ClipBottom = llngSH

    On Error Resume Next
    For X = 1 To T
        Graphs(GraphIndex).EvVal X, EvX, EvY, EvStat
        If EvStat <> 1 Then 'EVAL_ERROR
            CX = llngX0 + EvX * plngDocumentScaleX
            CY = llngY0 - EvY * plngDocumentScaleY

            If PrevEvStat <> 1 Then 'EVAL_ERROR
                'max gap?
                If Abs(PrevEvY - EvY) <= M Then
                    'clip: Cohen-Sutherland line-clipping algorithm, optimized by me
                    CSCodeL = CohenSutherlandComputeCode(LX, LY, ClipRight, ClipBottom)
                    CSCodeC = CohenSutherlandComputeCode(CX, CY, ClipRight, ClipBottom)

                    'both ends in the viewable area, no clipping needed
                    If (CSCodeL Or CSCodeC) = 0 Then
                        Line (LX, LY)-(CX, CY), C
                        'If (CX < 0) Or (CX > ClipRight) Or (CY < 0) Or (CY > ClipBottom) Then Me.Caption = X & " " & CX & " " & CY
                        GoTo SkipToNext
                    End If

                    SLX = LX: SLY = LY: SCX = CX: SCY = CY
                    Do
                        'trivial reject: both ends on the external side of the rectangle
                        If CSCodeL And CSCodeC Then GoTo SkipToNext

                        'normal case: clip end outside rectangle
                        If CSCodeL Then
                            CSCode = CSCodeL
                        Else
                            CSCode = CSCodeC
                        End If

                        If (CSCode And 1) Then 'top
                            SX = SLX + (SCX - SLX) * (ClipBottom - SLY) / (SCY - SLY)
                            SY = ClipBottom
                        ElseIf (CSCode And 2) Then 'bottom
                            'SX = SLX + (SCX - SLX) * (ClipTop - SLY) / (SCY - SLY)
                            'SY = ClipTop
                            SX = SLX + (SCX - SLX) * (-SLY) / (SCY - SLY)
                            SY = 0
                        ElseIf (CSCode And 4) Then 'right
                            SX = ClipRight
                            SY = SLY + (SCY - SLY) * (ClipRight - SLX) / (SCX - SLX)
                        Else 'left
                            'SX = ClipLeft
                            'SY = SLY + (SCY - SLY) * (ClipLeft - SLX) / (SCX - SLX)
                            SX = 0
                            SY = SLY + (SCY - SLY) * (-SLX) / (SCX - SLX)
                        End If

                        'set new end point and iterate
                        If CSCode = CSCodeL Then
                            SLX = SX: SLY = SY
                            CSCodeL = CohenSutherlandComputeCode(SLX, SLY, ClipRight, ClipBottom)
                        Else
                            SCX = SX: SCY = SY
                            CSCodeC = CohenSutherlandComputeCode(SCX, SCY, ClipRight, ClipBottom)
                        End If

                        'trivial accept: both ends in rectangle
                        If (CSCodeL Or CSCodeC) = 0 Then
                            Line (SLX, SLY)-(SCX, SCY), C
                            'If (SCX < 0) Or (SCX > ClipRight) Or (SCY < 0) Or (SCY > ClipBottom) Then Me.Caption = X & " " & SCX & " " & SCY
                            GoTo SkipToNext
                        End If
                    Loop
                End If
            End If

SkipToNext:
            LX = CX: LY = CY
            PrevEvY = EvY 'needed for MaxGap calculation
        End If

        PrevEvStat = EvStat
    Next X
End Sub

Public Sub WindowRefresh()
    If Me.WindowState = vbMinimized Then Exit Sub

    Me.Cls
    Me.BackColor = plngBackgroundColor

    If pblnGridVisible Then DrawGrid
    If pblnAxesVisible Then DrawAxes
    If pblnValuesVisible Then DrawValues

    'drawing the graphs
    'Using the form level lngTemp variable for temporary calculations
    For lngTemp = 1 To Graphs.Count
        If Graphs(lngTemp).Visible Then
            If Graphs(lngTemp).DrawLines Then
                GraphDrawLines lngTemp
            End If
            If Graphs(lngTemp).DrawPoints Then
                GraphDrawPoints lngTemp
            End If
        End If
    Next lngTemp

    Me.Refresh
End Sub

Public Sub FileNew()
    glngNewWindowNum = glngNewWindowNum + 1
    FileName = "Document" & LngToStr(glngNewWindowNum) & ".fng"
    Me.Caption = ShortenTitle(FileName)

    'finally
    ResizeEnabled = True
    WindowRefresh
End Sub

Public Sub FileOpen(OpenFilename As String)
    Me.FileName = OpenFilename
    Me.FileSaved = True
    'Me.FileChanged = False
    Me.Caption = ShortenTitle(GetFileTitle(OpenFilename))

    'parsing file
    If Not XMLFileParse Then
        Unload Me
        Exit Sub
    End If

    Dim I As Long, blnGraphsListChanged As Boolean
    For I = 1 To Graphs.Count
        'is the function valid?
        If Graphs(I).Expression.Build Then
            Graphs.Remove I
            blnGraphsListChanged = True
        Else
            Graphs(I).CacheValues
        End If
    Next I

    'if any of the graphs was removed because of a syntax error - update the list
    If blnGraphsListChanged Then frmMain.UpdateGraphsList

    'finally
    ResizeEnabled = True
    WindowRefresh
End Sub

Public Sub FileSave()
    'Print #1, "<?xml version=""1.0""?>"
    Print #1, "<DocumentScalingX>" & LngToStr(plngDocumentScaleX \ 15) & "</DocumentScalingX>"
    Print #1, "<DocumentScalingY>" & LngToStr(plngDocumentScaleY \ 15) & "</DocumentScalingY>"
    Print #1, "<OriginX>" & LngToStr(llngX0 \ Screen.TwipsPerPixelX) & "</OriginX>"
    Print #1, "<OriginY>" & LngToStr(llngY0 \ Screen.TwipsPerPixelY) & "</OriginY>"

    Print #1, "<BackgroundColor>" & ColorToStr(plngBackgroundColor) & "</BackgroundColor>"
    Print #1, "<Axes>"
    Print #1, vbTab & "<Visible>" & BlnToStr(pblnAxesVisible) & "</Visible>"
    Print #1, vbTab & "<Color>" & ColorToStr(plngAxesColor) & "</Color>"
    Print #1, vbTab & "<Style>" & LngToStr(plngAxesStyle) & "</Style>"
    Print #1, vbTab & "<Width>" & LngToStr(plngAxesWidth) & "</Width>"
    Print #1, "</Axes>"
    Print #1, "<Grid>"
    Print #1, vbTab & "<Visible>" & BlnToStr(pblnGridVisible) & "</Visible>"
    Print #1, vbTab & "<Color>" & ColorToStr(plngGridColor) & "</Color>"
    Print #1, vbTab & "<Style>" & LngToStr(plngGridStyle) & "</Style>"
    Print #1, vbTab & "<Width>" & LngToStr(plngGridWidth) & "</Width>"
    Print #1, "</Grid>"
    Print #1, "<Values>"
    Print #1, vbTab & "<Visible>" & BlnToStr(pblnValuesVisible) & "</Visible>"
    Print #1, vbTab & "<Color>" & ColorToStr(plngValuesColor) & "</Color>"
    Print #1, vbTab & "<Font>"
    Print #1, vbTab & vbTab & "<Name>" & pstrValuesFontName & "</Name>"
    Print #1, vbTab & vbTab & "<Size>" & LngToStr(plngValuesFontSize) & "</Size>"
    Print #1, vbTab & vbTab & "<Bold>" & BlnToStr(pblnValuesFontBold) & "</Bold>"
    Print #1, vbTab & vbTab & "<Italic>" & BlnToStr(pblnValuesFontBold) & "</Italic>"
    Print #1, vbTab & "</Font>"
    Print #1, "</Values>"

    Dim I As Long
    For I = 1 To Graphs.Count
        'Print #1,
        Print #1, "<Graph>"
        Print #1, vbTab & "<Function>" & Graphs(I).Expression.Expression & "</Function>"
        If Graphs(I).Description <> "" Then
            Print #1, vbTab & "<Description>" & Graphs(I).Description & "</Description>"
        Else
            Print #1, vbTab & "<Description />"
        End If
        Print #1, vbTab & "<Min>" & CurToStr(Graphs(I).Min) & "</Min>"
        Print #1, vbTab & "<Max>" & CurToStr(Graphs(I).Max) & "</Max>"
        Print #1, vbTab & "<Precision>" & CurToStr(Graphs(I).Precision) & "</Precision>"
        Print #1, vbTab & "<MaxGap>" & LngToStr(Graphs(I).MaxGap) & "</MaxGap>"
        Print #1, vbTab & "<DrawLines>" & BlnToStr(Graphs(I).DrawLines) & "</DrawLines>"
        Print #1, vbTab & "<DrawPoints>" & BlnToStr(Graphs(I).DrawPoints) & "</DrawPoints>"
        Print #1, vbTab & "<LinesWidth>" & LngToStr(Graphs(I).LinesWidth) & "</LinesWidth>"
        Print #1, vbTab & "<PointsSize>" & LngToStr(Graphs(I).PointsSize) & "</PointsSize>"
        Print #1, vbTab & "<Color>" & ColorToStr(Graphs(I).Color) & "</Color>"
        Print #1, vbTab & "<Visible>" & BlnToStr(Graphs(I).Visible) & "</Visible>"
        Print #1, "</Graph>"
    Next I
    For I = 1 To Variables.VariablesCollection.Count
        'Print #1,
        Print #1, "<Variable>"
        Print #1, vbTab & "<Name>" & Variables.VariablesCollection(I).Name & "</Name>"
        Print #1, vbTab & "<Value>" & CurToStr(Variables.VariablesCollection(I).Value) & "</Value>"
        Print #1, "</Variable>"
    Next I

    Me.FileSaved = True
    Me.FileChanged = False
End Sub

Private Function XMLFileParse() As Boolean
    Dim StartTagS As Boolean, EndTagS As Boolean
    Dim TagName As String, TagContent As String
    Dim Ch As String * 1, T As Long

    Do While Not EOF(1)
        Get #1, , Ch

        Select Case Ch
            Case "<"
                If Not StartTagS Then
                    StartTagS = True
                Else
                    EndTagS = True
                End If
            Case ">"
                If StartTagS Then
                    StartTagS = False
                Else
                    EndTagS = False
                End If

                If Left(TagName, 1) = "/" Then
                    'it is a close tag. strip leading "/"
                    TagName = Right(TagName, Len(TagName) - 1)
                    If TagName <> TagStack(TagDepth) Then
                        MsgBox "End tag '" & TagName & "' does not match the start tag '" & TagStack(TagDepth) & "'.", vbCritical
                        Exit Function
                    Else
                        XMLHandleTag TagName, TagContent, False
                        'TagStack(TagDepth) = ""
                        TagDepth = TagDepth - 1
                    End If
                Else
                    'it is an open tag
                    TagContent = ""
                    If Right(TagName, 1) <> "/" Then
                    'not an empty tag: <TagName/>
                        If TagDepth = TagDepthMax Then
                            MsgBox "Internal error: Max tag stack depth", vbCritical
                            Exit Function
                        Else
                            T = XMLFindWhiteSpace(TagName)
                            If T > 0 Then
                                'ignore tag attributes
                                TagName = Left(TagName, T - 1)
                            End If
                            XMLHandleTag TagName, TagContent, True

                            TagDepth = TagDepth + 1
                            TagStack(TagDepth) = TagName
                        End If
                    End If
                End If
                TagName = ""

            Case Else
                If StartTagS Or EndTagS Then
                    TagName = TagName + LCase(Ch)
                Else
                    TagContent = TagContent + Ch
                End If
        End Select
    Loop

    If TagDepth > 1 Then
        MsgBox "The '" & TagStack(TagDepth) & "' tag was not closed.", vbCritical
        Exit Function
    End If

    'file parsed successfully
    XMLFileParse = True
End Function

Private Sub XMLHandleTag(TagName As String, TagContent As String, StartTag As Boolean)
    Select Case TagName
        Case "documentscalingx"
            If Not StartTag Then
                If CheckTag(0, TagName, "", StartTag) Then
                    plngDocumentScaleX = StrToLng(TagContent) * 15
                End If
            End If

        Case "documentscalingy"
            If Not StartTag Then
                If CheckTag(0, TagName, "", StartTag) Then
                    plngDocumentScaleY = StrToLng(TagContent) * 15
                End If
            End If

        Case "originx"
            If Not StartTag Then
                If CheckTag(0, TagName, "", StartTag) Then
                    llngX0 = StrToLng(TagContent) * Screen.TwipsPerPixelX
                End If
            End If

        Case "originy"
            If Not StartTag Then
                If CheckTag(0, TagName, "", StartTag) Then
                    llngY0 = StrToLng(TagContent) * Screen.TwipsPerPixelY
                End If
            End If

        Case "backgroundcolor"
            If Not StartTag Then
                If CheckTag(0, TagName, "", StartTag) Then
                    plngBackgroundColor = StrToColor(TagContent)
                End If
            End If

        Case FNG_TAG_GRAPH
            If CheckTag(0, TagName, "", StartTag) Then
                If StartTag Then
                    GraphsAdd
                Else
                    'was at least the function set?
                    If Graphs(Graphs.Count).Expression.Expression = "" Then
                        Graphs.Remove Graphs.Count
                    End If
                End If
            End If

        Case "function"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_GRAPH, StartTag) Then
                    Graphs(Graphs.Count).Expression.Expression = TagContent
                End If
            End If

        Case "description"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_GRAPH, StartTag) Then
                    Graphs(Graphs.Count).Description = TagContent
                End If
            End If

        Case "min"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_GRAPH, StartTag) Then
                    Graphs(Graphs.Count).Min = StrToLng(TagContent)
                End If
            End If

        Case "max"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_GRAPH, StartTag) Then
                    Graphs(Graphs.Count).Max = StrToLng(TagContent)
                End If
            End If

        Case "precision"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_GRAPH, StartTag) Then
                    Graphs(Graphs.Count).Precision = StrToCur(TagContent)
                End If
            End If

        Case "maxgap"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_GRAPH, StartTag) Then
                    Graphs(Graphs.Count).MaxGap = StrToLng(TagContent)
                End If
            End If

        Case "drawlines"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_GRAPH, StartTag) Then
                    Graphs(Graphs.Count).DrawLines = StrToBln(TagContent)
                End If
            End If

        Case "drawpoints"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_GRAPH, StartTag) Then
                    Graphs(Graphs.Count).DrawPoints = StrToBln(TagContent)
                End If
            End If

        Case "lineswidth"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_GRAPH, StartTag) Then
                    Graphs(Graphs.Count).LinesWidth = StrToLng(TagContent)
                End If
            End If

        Case "pointssize"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_GRAPH, StartTag) Then
                    Graphs(Graphs.Count).PointsSize = StrToLng(TagContent)
                End If
            End If

        Case "color"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_GRAPH, StartTag) Then
                    Graphs(Graphs.Count).Color = StrToColor(TagContent)
                End If
                If CheckTag(1, TagName, FNG_TAG_AXES, StartTag) Then
                    plngAxesColor = StrToColor(TagContent)
                End If
                If CheckTag(1, TagName, FNG_TAG_GRID, StartTag) Then
                    plngGridColor = StrToColor(TagContent)
                End If
                If CheckTag(1, TagName, FNG_TAG_VALUES, StartTag) Then
                    plngValuesColor = StrToColor(TagContent)
                End If
            End If

        Case "visible"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_GRAPH, StartTag) Then
                    Graphs(Graphs.Count).Visible = StrToBln(TagContent)
                End If
                If CheckTag(1, TagName, FNG_TAG_AXES, StartTag) Then
                    pblnAxesVisible = StrToBln(TagContent)
                End If
                If CheckTag(1, TagName, FNG_TAG_GRID, StartTag) Then
                    pblnGridVisible = StrToBln(TagContent)
                End If
                If CheckTag(1, TagName, FNG_TAG_VALUES, StartTag) Then
                    pblnValuesVisible = StrToBln(TagContent)
                End If
            End If

        Case "style"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_AXES, StartTag) Then
                    plngAxesStyle = StrToLng(TagContent)
                End If
                If CheckTag(1, TagName, FNG_TAG_GRID, StartTag) Then
                    plngGridStyle = StrToLng(TagContent)
                End If
            End If

        Case "width"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_AXES, StartTag) Then
                    plngAxesWidth = StrToLng(TagContent)
                End If
                If CheckTag(1, TagName, FNG_TAG_GRID, StartTag) Then
                    plngGridWidth = StrToLng(TagContent)
                End If
            End If

        Case FNG_TAG_VARIABLE
            If CheckTag(0, TagName, "", StartTag) Then
                If StartTag Then
                    Dim NewVariable As New clsVariable
                    Variables.VariablesCollection.Add NewVariable
                Else
                    If Variables.VariablesCollection(Variables.VariablesCollection.Count).Name = "" Then
                        'no name was set, delete the variable
                        Variables.VariablesCollection.Remove Variables.VariablesCollection.Count
                    End If
                End If
            End If
            
        Case "name"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_VARIABLE, StartTag) Then
                    Variables.VariablesCollection(Variables.VariablesCollection.Count).Name = TagContent
                End If
            End If

        Case "value"
            If Not StartTag Then
                If CheckTag(1, TagName, FNG_TAG_VARIABLE, StartTag) Then
                    Variables.VariablesCollection(Variables.VariablesCollection.Count).Value = StrToCur(TagContent)
                End If
            End If
    End Select
End Sub

Private Function CheckTag(BaseTagDepth As Long, TagName As String, ParentTagName As String, StartTag As Boolean) As Boolean
    'Returns True if check passed, otherwise returns False

    If StartTag Then
        If TagDepth <> BaseTagDepth Then Exit Function
    Else
        If TagDepth <> BaseTagDepth + 1 Then Exit Function
        If BaseTagDepth > 0 Then
            If TagStack(BaseTagDepth) <> ParentTagName Then Exit Function
        End If
        If TagStack(BaseTagDepth + 1) <> TagName Then Exit Function
    End If

    CheckTag = True
End Function
