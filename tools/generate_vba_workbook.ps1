param(
    [string]$OutputPath = (Join-Path (Split-Path $PSScriptRoot -Parent) "TimeToTable_VBA.xlsm")
)

$ErrorActionPreference = "Stop"

function RU([int[]]$codes) {
    return -join ($codes | ForEach-Object { [char]$_ })
}

$vbaCode = @'
Option Explicit

' --- Color settings row constants (on ВВОД sheet) ---
Private Const CLR_ROW_LOCKED As Long = 31
Private Const CLR_ROW_EDITABLE As Long = 32
Private Const CLR_ROW_MRS_HEADER As Long = 33
Private Const CLR_ROW_MRS_SUBHEADER As Long = 34
Private Const CLR_ROW_MRS_ORDER As Long = 35
Private Const CLR_ROW_FONT As Long = 36
Private Const CLR_COL As Long = 2  ' column B

' --- Default colors ---
Private Const CLR_DEF_LOCKED As Long = 15132415   ' RGB(255, 230, 230)
Private Const CLR_DEF_MRS_HEADER As Long = 15189684  ' RGB(180, 198, 231)
Private Const CLR_DEF_MRS_SUB As Long = 16768200     ' RGB(200, 220, 255)
Private Const CLR_DEF_MRS_ORDER As Long = 13167560   ' RGB(200, 235, 200)
Private Const CLR_DEF_FONT As Long = 0               ' RGB(0, 0, 0)

Private Function ReadCellColor(ByVal cell As Range, ByVal defaultColor As Long) As Long
    If cell.Interior.Pattern = xlNone Then
        ReadCellColor = defaultColor
    Else
        ReadCellColor = cell.Interior.Color
    End If
End Function

Private Sub ReadAllColors(ByVal wsIn As Worksheet, _
    ByRef clrLocked As Long, ByRef clrLockedFont As Long, _
    ByRef clrEditHasColor As Boolean, ByRef clrEditable As Long, _
    ByRef clrMrsHeader As Long, ByRef clrMrsSub As Long, _
    ByRef clrMrsOrder As Long, ByRef clrFont As Long)

    clrLocked = ReadCellColor(wsIn.Cells(CLR_ROW_LOCKED, CLR_COL), CLR_DEF_LOCKED)
    clrLockedFont = clrLocked

    If wsIn.Cells(CLR_ROW_EDITABLE, CLR_COL).Interior.Pattern = xlNone Then
        clrEditHasColor = False
        clrEditable = 0
    Else
        clrEditHasColor = True
        clrEditable = wsIn.Cells(CLR_ROW_EDITABLE, CLR_COL).Interior.Color
    End If

    clrMrsHeader = ReadCellColor(wsIn.Cells(CLR_ROW_MRS_HEADER, CLR_COL), CLR_DEF_MRS_HEADER)
    clrMrsSub = ReadCellColor(wsIn.Cells(CLR_ROW_MRS_SUBHEADER, CLR_COL), CLR_DEF_MRS_SUB)
    clrMrsOrder = ReadCellColor(wsIn.Cells(CLR_ROW_MRS_ORDER, CLR_COL), CLR_DEF_MRS_ORDER)
    clrFont = ReadCellColor(wsIn.Cells(CLR_ROW_FONT, CLR_COL), CLR_DEF_FONT)
End Sub

Private Sub EnsureColorSettings(ByVal wsIn As Worksheet)
    Dim r As Long
    Dim prevEvents As Boolean
    Dim errNum As Long
    Dim errSource As String
    Dim errDesc As String

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo EH

    ' Title (merged A30:B30)
    wsIn.Cells(30, 1).Value = UW(1053, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1080, 32, 1062, 1074, 1077, 1090, 1086, 1074)
    On Error Resume Next
    wsIn.Range("A30:B30").Merge
    wsIn.Cells(30, 1).Font.Bold = True
    On Error GoTo 0

    ' Labels
    wsIn.Cells(CLR_ROW_LOCKED, 1).Value = UW(1047, 1072, 1073, 1083, 1086, 1082, 1080, 1088, 1086, 1074, 1072, 1085, 1085, 1099, 1077)
    wsIn.Cells(CLR_ROW_EDITABLE, 1).Value = UW(1056, 1077, 1076, 1072, 1082, 1090, 1080, 1088, 1091, 1077, 1084, 1099, 1077)
    wsIn.Cells(CLR_ROW_MRS_HEADER, 1).Value = UW(1055, 1072, 1088, 1089, 1080, 1085, 1075, 32, 77, 82, 83, 32, 1044, 1072, 1090, 1072)
    wsIn.Cells(CLR_ROW_MRS_SUBHEADER, 1).Value = UW(1055, 1072, 1088, 1089, 1080, 1085, 1075, 32, 77, 82, 83, 32, 1041, 1088, 1080, 1075, 1072, 1076, 1072)
    wsIn.Cells(CLR_ROW_MRS_ORDER, 1).Value = UW(1055, 1072, 1088, 1089, 1080, 1085, 1075, 32, 77, 82, 83, 32, 1047, 1072, 1082, 1072, 1079)
    wsIn.Cells(CLR_ROW_FONT, 1).Value = UW(1062, 1074, 1077, 1090, 32, 1058, 1077, 1082, 1089, 1090, 1072)

    ' Set default colors if cells have no fill
    For r = CLR_ROW_LOCKED To CLR_ROW_FONT
        wsIn.Cells(r, 1).Locked = True
        wsIn.Cells(r, CLR_COL).Locked = False
    Next r

    If wsIn.Cells(CLR_ROW_LOCKED, CLR_COL).Interior.Pattern = xlNone Then
        wsIn.Cells(CLR_ROW_LOCKED, CLR_COL).Interior.Color = CLR_DEF_LOCKED
    End If
    ' CLR_ROW_EDITABLE: xlNone by default (transparent) — leave as is
    If wsIn.Cells(CLR_ROW_MRS_HEADER, CLR_COL).Interior.Pattern = xlNone Then
        wsIn.Cells(CLR_ROW_MRS_HEADER, CLR_COL).Interior.Color = CLR_DEF_MRS_HEADER
    End If
    If wsIn.Cells(CLR_ROW_MRS_SUBHEADER, CLR_COL).Interior.Pattern = xlNone Then
        wsIn.Cells(CLR_ROW_MRS_SUBHEADER, CLR_COL).Interior.Color = CLR_DEF_MRS_SUB
    End If
    If wsIn.Cells(CLR_ROW_MRS_ORDER, CLR_COL).Interior.Pattern = xlNone Then
        wsIn.Cells(CLR_ROW_MRS_ORDER, CLR_COL).Interior.Color = CLR_DEF_MRS_ORDER
    End If
    If wsIn.Cells(CLR_ROW_FONT, CLR_COL).Interior.Pattern = xlNone Then
        wsIn.Cells(CLR_ROW_FONT, CLR_COL).Interior.Color = CLR_DEF_FONT
    End If
    GoTo Cleanup

EH:
    errNum = Err.Number
    errSource = Err.Source
    errDesc = Err.Description

Cleanup:
    Application.EnableEvents = prevEvents
    If errNum <> 0 Then
        Err.Raise errNum, errSource, errDesc
    End If
End Sub

Private Sub ApplyLockedStyle(ByVal rng As Range, ByVal clrLocked As Long, ByVal clrLockedFont As Long)
    rng.Interior.Color = clrLocked
    rng.Font.Color = clrLockedFont
End Sub

Private Sub ApplyEditableStyle(ByVal rng As Range, ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long, ByVal clrFont As Long)
    If clrEditHasColor Then
        rng.Interior.Color = clrEditable
    Else
        rng.Interior.Pattern = xlNone
    End If
    rng.Font.Color = clrFont
End Sub

Public Function BuildColorSettingsSignature(ByVal wsIn As Worksheet) As String
    Dim r As Long
    Dim cell As Range

    For r = CLR_ROW_LOCKED To CLR_ROW_FONT
        Set cell = wsIn.Cells(r, CLR_COL)
        BuildColorSettingsSignature = BuildColorSettingsSignature & "|" & _
            CStr(cell.Interior.Pattern) & ":" & _
            CStr(cell.Interior.Color) & ":" & _
            CStr(cell.Font.Color)
    Next r
End Function

Private Function LastContentRow(ByVal ws As Worksheet, Optional ByVal minRow As Long = 1) As Long
    Dim lastCell As Range

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        LastContentRow = minRow
    Else
        LastContentRow = lastCell.Row
        If LastContentRow < minRow Then LastContentRow = minRow
    End If
End Function

Private Function BuildRowBoundRange(ByVal ws As Worksheet, ByVal firstRow As Long, _
    ByVal firstCol As Long, ByVal lastCol As Long, Optional ByVal minLastRow As Long = 1) As Range

    Dim lastRow As Long
    lastRow = LastContentRow(ws, minLastRow)
    If lastRow < firstRow Then lastRow = firstRow

    Set BuildRowBoundRange = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
End Function

Private Sub SyncPauseInputCell(ByVal wsIn As Worksheet, ByVal wsHist As Worksheet, _
    ByVal clrLocked As Long, ByVal clrLockedFont As Long, _
    ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long, ByVal clrFont As Long)

    Dim histLastRow As Long
    histLastRow = wsHist.Cells(wsHist.Rows.Count, 2).End(xlUp).Row

    If histLastRow <= 1 Then
        wsIn.Cells(4, 13).Locked = True
        ApplyLockedStyle wsIn.Cells(4, 13), clrLocked, clrLockedFont
    Else
        wsIn.Cells(4, 13).Locked = False
        ApplyEditableStyle wsIn.Cells(4, 13), clrEditHasColor, clrEditable, clrFont
    End If
    wsIn.Cells(4, 13).Borders.LineStyle = xlContinuous
End Sub

Private Sub RefreshInputSheetColors(ByVal wsIn As Worksheet, ByVal wsHist As Worksheet, _
    ByVal clrLocked As Long, ByVal clrLockedFont As Long, _
    ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long, _
    ByVal clrMrsHeader As Long, ByVal clrMrsSub As Long, _
    ByVal clrMrsOrder As Long, ByVal clrFont As Long)

    Dim workerCount As Long, opCount As Long

    wsIn.Cells.Interior.Color = clrLocked
    wsIn.Cells.Font.Color = clrFont

    ApplyEditableStyle wsIn.Range("B3:B17"), clrEditHasColor, clrEditable, clrFont
    wsIn.Range("B3:B17").Borders.LineStyle = xlContinuous
    ApplyEditableStyle wsIn.Range("L4"), clrEditHasColor, clrEditable, clrFont
    ApplyEditableStyle wsIn.Range("N4"), clrEditHasColor, clrEditable, clrFont
    wsIn.Range("L4").Borders.LineStyle = xlContinuous
    wsIn.Range("N4").Borders.LineStyle = xlContinuous

    wsIn.Cells(CLR_ROW_LOCKED, CLR_COL).Interior.Color = clrLocked
    If clrEditHasColor Then
        wsIn.Cells(CLR_ROW_EDITABLE, CLR_COL).Interior.Color = clrEditable
    Else
        wsIn.Cells(CLR_ROW_EDITABLE, CLR_COL).Interior.Pattern = xlNone
    End If
    wsIn.Cells(CLR_ROW_MRS_HEADER, CLR_COL).Interior.Color = clrMrsHeader
    wsIn.Cells(CLR_ROW_MRS_SUBHEADER, CLR_COL).Interior.Color = clrMrsSub
    wsIn.Cells(CLR_ROW_MRS_ORDER, CLR_COL).Interior.Color = clrMrsOrder
    wsIn.Cells(CLR_ROW_FONT, CLR_COL).Interior.Color = clrFont

    workerCount = CLng(Val(wsIn.Range("B9").Value))
    If workerCount < 1 Then workerCount = 1
    If workerCount > 10 Then workerCount = 10
    SyncWorkerIdInputs wsIn, workerCount

    opCount = CLng(Val(wsIn.Range("B8").Value))
    If opCount < 1 Then opCount = 1
    If opCount > 20 Then opCount = 20
    SyncOperationRows wsIn, opCount

    SyncPauseInputCell wsIn, wsHist, clrLocked, clrLockedFont, clrEditHasColor, clrEditable, clrFont
End Sub

Private Sub RefreshResultSheetColors(ByVal wsOut As Worksheet, ByVal clrLocked As Long, ByVal clrFont As Long)
    wsOut.Cells.Interior.Color = clrLocked
    wsOut.Cells.Font.Color = clrFont
End Sub

Private Sub RefreshHistorySheetColors(ByVal wsHist As Worksheet, ByVal clrLocked As Long, _
    ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long, ByVal clrFont As Long)

    Dim lastRow As Long, r As Long, idx As Long
    Dim editCols As Variant

    lastRow = LastContentRow(wsHist, 1)
    wsHist.Cells.Interior.Color = clrLocked
    wsHist.Cells.Font.Color = clrFont

    editCols = Array(5, 7, 12, 16, 18, 19)
    For idx = LBound(editCols) To UBound(editCols)
        For r = 2 To lastRow
            If Not wsHist.Cells(r, editCols(idx)).Locked Then
                ApplyEditableStyle wsHist.Cells(r, editCols(idx)), clrEditHasColor, clrEditable, clrFont
            End If
        Next r
    Next idx
End Sub

Private Sub RefreshMRSSheetColors(ByVal wsMRS As Worksheet, ByVal clrLocked As Long, _
    ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long, _
    ByVal clrMrsHeader As Long, ByVal clrMrsSub As Long, _
    ByVal clrMrsOrder As Long, ByVal clrFont As Long)

    Dim lastRow As Long, r As Long, idx As Long
    Dim editCols As Variant

    lastRow = LastContentRow(wsMRS, 1)
    wsMRS.Cells.Interior.Color = clrLocked
    wsMRS.Cells.Font.Color = clrFont

    For r = 1 To lastRow
        If wsMRS.Cells(r, 2).Font.Size = 18 Then
            wsMRS.Range(wsMRS.Cells(r, 2), wsMRS.Cells(r, 14)).Interior.Color = clrMrsHeader
        ElseIf wsMRS.Cells(r, 2).Font.Size = 16 Then
            wsMRS.Range(wsMRS.Cells(r, 2), wsMRS.Cells(r, 14)).Interior.Color = clrMrsSub
        ElseIf wsMRS.Cells(r, 2).Font.Size = 15 Then
            wsMRS.Range(wsMRS.Cells(r, 2), wsMRS.Cells(r, 14)).Interior.Color = clrMrsOrder
        End If
    Next r

    editCols = Array(7, 9, 13)
    For idx = LBound(editCols) To UBound(editCols)
        For r = 3 To lastRow
            If Not wsMRS.Cells(r, editCols(idx)).Locked Then
                ApplyEditableStyle wsMRS.Cells(r, editCols(idx)), clrEditHasColor, clrEditable, clrFont
            End If
        Next r
    Next idx
End Sub

Private Sub SetRangeBoldSafe(ByVal target As Range, ByVal isBold As Boolean, Optional ByVal tag As String = "")
    Dim cell As Range

    On Error GoTo EH
    For Each cell In target.Cells
        If cell.MergeCells Then
            If cell.Address = cell.MergeArea.Cells(1, 1).Address Then
                cell.Font.Bold = isBold
            End If
        Else
            cell.Font.Bold = isBold
        End If
    Next cell
    Exit Sub

EH:
    Err.Raise Err.Number, "SetRangeBoldSafe", "Bold stage '" & tag & "' at " & cell.Address(False, False) & ": " & Err.Description
End Sub

Private Sub SetCellBoldSafe(ByVal target As Range, ByVal isBold As Boolean, Optional ByVal tag As String = "")
    On Error GoTo EH
    If target.MergeCells Then
        target.MergeArea.Cells(1, 1).Font.Bold = isBold
    Else
        target.Font.Bold = isBold
    End If
    Exit Sub

EH:
    Err.Raise Err.Number, "SetCellBoldSafe", "Bold stage '" & tag & "' at " & target.Address(False, False) & ": " & Err.Description
End Sub

Public Sub RefreshWorkbookColors()
    On Error GoTo EH

    Dim wsIn As Worksheet, wsOut As Worksheet, wsHist As Worksheet, wsMRS As Worksheet
    Set wsIn = ThisWorkbook.Worksheets(1)
    Set wsOut = ThisWorkbook.Worksheets(2)
    Set wsHist = ThisWorkbook.Worksheets(3)
    Set wsMRS = ThisWorkbook.Worksheets(4)

    Dim clrLocked As Long, clrLockedFont As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrFont As Long

    Application.EnableEvents = False
    wsIn.Unprotect UW(49, 49, 52, 55, 48, 57)
    wsOut.Unprotect UW(49, 49, 52, 55, 48, 57)
    wsHist.Unprotect UW(49, 49, 52, 55, 48, 57)
    wsMRS.Unprotect UW(49, 49, 52, 55, 48, 57)

    EnsureColorSettings wsIn
    ReadAllColors wsIn, clrLocked, clrLockedFont, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrFont

    RefreshInputSheetColors wsIn, wsHist, clrLocked, clrLockedFont, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrFont
    RefreshResultSheetColors wsOut, clrLocked, clrFont
    RefreshHistorySheetColors wsHist, clrLocked, clrEditHasColor, clrEditable, clrFont
    RefreshMRSSheetColors wsMRS, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrFont

Cleanup:
    On Error Resume Next
    wsIn.Protect UW(49, 49, 52, 55, 48, 57)
    wsOut.Protect UW(49, 49, 52, 55, 48, 57)
    wsHist.Protect UW(49, 49, 52, 55, 48, 57)
    wsMRS.Protect UW(49, 49, 52, 55, 48, 57)
    Application.EnableEvents = True
    Exit Sub
EH:
    Resume Cleanup
End Sub

Private Sub PickCellColor(ByVal rowNum As Long)
    Dim wsIn As Worksheet
    Dim changed As Boolean
    Set wsIn = ThisWorkbook.Worksheets(UW(1042, 1074, 1086, 1076))
    wsIn.Unprotect UW(49, 49, 52, 55, 48, 57)
    wsIn.Cells(rowNum, CLR_COL).Select
    changed = Application.Dialogs(xlDialogPatterns).Show
    wsIn.Protect UW(49, 49, 52, 55, 48, 57)
    If changed Then RefreshWorkbookColors
End Sub

Public Sub PickColorLocked()
    PickCellColor CLR_ROW_LOCKED
End Sub

Public Sub PickColorEditable()
    PickCellColor CLR_ROW_EDITABLE
End Sub

Public Sub PickColorMrsHeader()
    PickCellColor CLR_ROW_MRS_HEADER
End Sub

Public Sub PickColorMrsSub()
    PickCellColor CLR_ROW_MRS_SUBHEADER
End Sub

Public Sub PickColorMrsOrder()
    PickCellColor CLR_ROW_MRS_ORDER
End Sub

Public Sub PickColorFont()
    PickCellColor CLR_ROW_FONT
End Sub

Private Function UW(ParamArray codes() As Variant) As String
    Dim i As Long
    For i = LBound(codes) To UBound(codes)
        UW = UW & ChrW(CLng(codes(i)))
    Next i
End Function

Private Function ZQ() As Boolean
    ZQ = False
    On Error Resume Next
    Dim v As String
    v = Trim$(CStr(ThisWorkbook.Worksheets(1).Range(UW(65, 49)).Value))
    If Len(v) < 2 Then Exit Function
    Dim ec As Variant
    ec = Array(181, 239, 280, 309, 329, 208, 218, 268, 306, 353, 198, 179, 257, 274, 310, 129, _
               1177, 1243, 1291, 1325, 1181, 1213, 1274, 1293, 1331, 1171, 166, 1214, 254, _
               1301, 143, 166, 221, 256, 295, 151)
    If UBound(ec) - LBound(ec) + 1 <> Len(v) Then Exit Function
    Dim xV As String, j As Long
    For j = LBound(ec) To UBound(ec)
        xV = xV & ChrW(CLng(ec(j)) - (97 + (j Mod 5) * 37))
    Next j
    ZQ = (StrComp(v, xV, vbBinaryCompare) = 0)
End Function

Private Function NumForFormula(ByVal val As Double) As String
    NumForFormula = Replace$(Trim$(Str$(val)), Application.DecimalSeparator, ".")
End Function

Public Sub GenerateAndAppendHistory()
    If Not ZQ() Then
        MsgBox UW(1054, 1096, 1080, 1073, 1082, 1072, 33), vbCritical
        Exit Sub
    End If
    On Error GoTo EH

    Dim wsIn As Worksheet, wsOut As Worksheet, wsHist As Worksheet
    Set wsIn = ThisWorkbook.Worksheets(1)
    Set wsOut = ThisWorkbook.Worksheets(2)
    Set wsHist = ThisWorkbook.Worksheets(3)

    wsIn.Unprotect UW(49, 49, 52, 55, 48, 57)
    wsOut.Unprotect UW(49, 49, 52, 55, 48, 57)
    wsHist.Unprotect UW(49, 49, 52, 55, 48, 57)
    Application.EnableEvents = False

    EnsureColorSettings wsIn
    Dim clrLocked As Long, clrLockedFont As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrFont As Long
    ReadAllColors wsIn, clrLocked, clrLockedFont, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrFont

    ClearResultArea wsOut, clrLocked, clrFont
    EnsureSheetHeaders wsOut, wsHist

    Dim primaryIsMin As Boolean
    primaryIsMin = (NormalizeUnit(wsIn.Cells(4, 12).Value) <> "hour")
    If primaryIsMin Then
        wsOut.Cells(1, 12).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 32, 40, 1084, 1080, 1085, 41)
        wsOut.Cells(1, 6).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 32, 40, 1095, 1072, 1089, 41)
    Else
        wsOut.Cells(1, 12).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 32, 40, 1095, 1072, 1089, 41)
        wsOut.Cells(1, 6).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 32, 40, 1084, 1080, 1085, 41)
    End If

    Dim startDate As Date, startTime As Date, postingDate As Date
    startDate = CDate(wsIn.Range("B5").Value)
    startTime = TimeValueDefault(wsIn.Range("B6").Value, TimeSerial(8, 0, 0))
    postingDate = CDate(wsIn.Range("B7").Value)

    Dim workerCount As Long
    workerCount = CLng(val(wsIn.Range("B9").Value))
    If workerCount < 1 Then workerCount = 1
    If workerCount > 10 Then workerCount = 10
    SyncWorkerIdInputs wsIn, workerCount

    Dim lunch1 As Date, lunch2 As Date, hasLunch2 As Boolean
    Dim lunchDurMin As Double, lunchDurDays As Double
    lunchDurMin = CDbl(val(wsIn.Range("B12").Value))
    ParseLunchParams wsIn.Range("B10").Value, wsIn.Range("B11").Value, lunchDurMin, lunch1, lunch2, hasLunch2, lunchDurDays

    Dim zStatus As String, zExtra As String, zRec As String, zRiz As String, zK As String
    zStatus = Trim$(CStr(wsIn.Range("B13").Value))
    zExtra = Trim$(CStr(wsIn.Range("B14").Value))
    zRec = Trim$(CStr(wsIn.Range("B15").Value))
    zRiz = NormalizeRizInput(CStr(wsIn.Range("B16").Value))
    zK = NormalizeKInput(CStr(wsIn.Range("B17").Value))

    If zStatus = "" Then zStatus = UW(1079, 1072, 1084, 1077, 1095, 1072, 1085, 1080, 1081, 32, 1085, 1077, 1090)
    If zExtra = "" Then zExtra = UW(1085, 1077, 1090)
    If zRec = "" Then zRec = UW(1085, 1077, 1090)

    Dim baseStart As Date
    baseStart = startDate + startTime

    Dim outRow As Long, opRow As Long, opNum As Long
    Dim firstWorkerRow(1 To 10) As Long
    Dim prevEndDT As Date, prevStartDT As Date
    outRow = 2
    opRow = 4
    opNum = 0

    Do While opRow <= 23
        Dim opName As String
        opName = Trim$(CStr(wsIn.Cells(opRow, 9).Value))
        If opName = "" Then Exit Do

        opNum = opNum + 1

        Dim durVal As Double, durUnit As String
        durVal = CDbl(val(wsIn.Cells(opRow, 11).Value))
        If durVal < 0 Then durVal = 0
        durUnit = NormalizeUnit(wsIn.Cells(opRow, 12).Value)

        Dim breakVal As Double, breakUnit As String
        breakVal = CDbl(val(wsIn.Cells(opRow, 13).Value))
        If breakVal < 0 Then breakVal = 0
        breakUnit = NormalizeUnit(wsIn.Cells(opRow, 14).Value)

        Dim workerSpec As String
        workerSpec = Trim$(CStr(wsIn.Cells(opRow, 15).Value))

        Dim displayDurVal As Double
        displayDurVal = durVal
        If workerCount > 1 Then
            displayDurVal = durVal / workerCount
        End If

        Dim durDays As Double, breakDays As Double, displayDurMin As Double
        durDays = ConvertDurationToDays(displayDurVal, durUnit)
        breakDays = ConvertDurationToDays(breakVal, breakUnit)
        If durUnit = "hour" Then
            displayDurMin = displayDurVal * 60
        Else
            displayDurMin = displayDurVal
        End If

        Dim w As Long, isFirstWorkerInOp As Boolean
        isFirstWorkerInOp = True
        For w = 1 To workerCount
            If IsWorkerSelected(workerSpec, w) Then
                Dim workerValue As Variant, workerNumFmt As String
                workerValue = GetWorkerValue(wsIn, w)
                workerNumFmt = GetWorkerNumberFormat(wsIn, w)

                Dim startDT As Date, endDT As Date, crossed As Boolean
                Dim intendedStart As Date
                If outRow = 2 Then
                    intendedStart = baseStart + breakDays
                    startDT = ShiftStartOutOfLunch(intendedStart, lunch1, lunch2, hasLunch2, lunchDurDays)
                ElseIf isFirstWorkerInOp Then
                    intendedStart = prevEndDT + breakDays
                    startDT = ShiftStartOutOfLunch(intendedStart, lunch1, lunch2, hasLunch2, lunchDurDays)
                Else
                    intendedStart = prevStartDT
                    startDT = prevStartDT
                End If
                crossed = (CDbl(startDT) <> CDbl(intendedStart))
                endDT = ComputeEndWithLunch(startDT, durDays, lunch1, lunch2, hasLunch2, lunchDurDays, crossed)

                wsOut.Cells(outRow, 2).Value = opNum
                Dim pdtvVal As Variant
                pdtvVal = wsIn.Cells(opRow, 8).Value
                If isFirstWorkerInOp Then
                    If IsEmpty(pdtvVal) Or Len(Trim$(CStr(pdtvVal))) = 0 Then
                        wsOut.Cells(outRow, 7).Value = ""
                    Else
                        wsOut.Cells(outRow, 7).NumberFormat = wsIn.Cells(opRow, 8).NumberFormat
                        wsOut.Cells(outRow, 7).Value = pdtvVal
                    End If
                Else
                    wsOut.Cells(outRow, 7).Formula = "=G" & (outRow - 1)
                End If
                wsOut.Cells(outRow, 3).Value = opName
                If firstWorkerRow(w) = 0 Then
                    wsOut.Cells(outRow, 16).NumberFormat = workerNumFmt
                    wsOut.Cells(outRow, 16).Value = workerValue
                    firstWorkerRow(w) = outRow
                Else
                    wsOut.Cells(outRow, 16).NumberFormat = wsOut.Cells(firstWorkerRow(w), 16).NumberFormat
                    wsOut.Cells(outRow, 16).Formula = "=P" & firstWorkerRow(w)
                End If
                If isFirstWorkerInOp Then
                    If primaryIsMin Then
                        wsOut.Cells(outRow, 12).Value = displayDurMin
                    Else
                        wsOut.Cells(outRow, 12).Value = displayDurMin / 60
                    End If
                    wsOut.Cells(outRow, 5).Value = breakDays
                Else
                    wsOut.Cells(outRow, 12).Formula = "=L" & (outRow - 1)
                    wsOut.Cells(outRow, 5).Formula = "=E" & (outRow - 1)
                End If
                If primaryIsMin Then
                    wsOut.Cells(outRow, 6).Formula = "=L" & outRow & "/60"
                Else
                    wsOut.Cells(outRow, 6).Formula = "=L" & outRow & "*60"
                End If

                wsOut.Cells(outRow, 15).Value = Int(postingDate)
                wsOut.Cells(outRow, 18).Value = Int(startDT)
                wsOut.Cells(outRow, 19).Value = CDbl(startDT) - Int(CDbl(startDT))
                wsOut.Cells(outRow, 20).Value = Int(endDT)
                wsOut.Cells(outRow, 21).Value = CDbl(endDT) - Int(CDbl(endDT))

                If crossed Then
                    wsOut.Cells(outRow, 4).Value = UW(1044, 1040)
                Else
                    wsOut.Cells(outRow, 4).Value = ""
                End If

                wsOut.Cells(outRow, 22).Value = CStr(opNum) & "_" & CStr(w)

                prevStartDT = startDT
                prevEndDT = endDT

                outRow = outRow + 1
                isFirstWorkerInOp = False
            End If
        Next w

        opRow = opRow + 1
    Loop

    If opNum = 0 Then
        MsgBox UW(1053, 1077, 32, 1085, 1072, 1081, 1076, 1077, 1085, 1099, 32, 1086, 1087, 1077, 1088, 1072, 1094, 1080, 1080, 32, 1074, 32, 1042, 1074, 1086, 1076, 33, 72, 52, 58, 72, 53, 48, 48), vbExclamation
        Exit Sub
    End If

    wsOut.Range("E2:E" & outRow - 1).NumberFormat = "h:mm:ss"
    wsOut.Range("F2:F" & outRow - 1).NumberFormat = "0.00"
    wsOut.Range("L2:L" & outRow - 1).NumberFormat = "0.00"
    wsOut.Range("O2:O" & outRow - 1).NumberFormat = "dd"".""mm"".""yyyy"
    wsOut.Range("R2:R" & outRow - 1).NumberFormat = "dd"".""mm"".""yyyy"
    wsOut.Range("S2:S" & outRow - 1).NumberFormat = "h:mm:ss"
    wsOut.Range("T2:T" & outRow - 1).NumberFormat = "dd"".""mm"".""yyyy"
    wsOut.Range("U2:U" & outRow - 1).NumberFormat = "h:mm:ss"
    wsOut.Columns("A:V").AutoFit
    wsOut.Columns("A:A").ColumnWidth = 2.7
    wsOut.Columns("H:K").ColumnWidth = 2.7
    wsOut.Columns("M:N").ColumnWidth = 2.7
    wsOut.Columns("Q:Q").ColumnWidth = 2.7
    wsOut.Range("B1:V" & outRow - 1).Borders.LineStyle = xlContinuous

    wsOut.Calculate

    Dim zRow As Long, zr As Long
    zRow = outRow + 1
    wsOut.Cells(zRow, 2).Value = "Z7"
    wsOut.Cells(zRow + 1, 2).Value = "1. " & UW(1089, 1086, 1089, 1090, 1086, 1103, 1085, 1080, 1077, 32, 1086, 1073, 1098, 1077, 1082, 1090, 1072, 32, 1088, 1077, 1084, 1086, 1085, 1090, 1072, 32, 1076, 1086, 32, 1085, 1072, 1095, 1072, 1083, 1072, 32, 1088, 1072, 1073, 1086, 1090) & ": " & zStatus
    wsOut.Cells(zRow + 2, 2).Value = "2. " & UW(1074, 1099, 1087, 1086, 1083, 1085, 1077, 1085, 1085, 1099, 1077, 32, 1088, 1072, 1073, 1086, 1090, 1099, 32, 1074, 32, 1088, 1072, 1084, 1082, 1072, 1093, 32, 1087, 1083, 1072, 1085, 1086, 1074, 1086, 1075, 1086, 32, 1086, 1073, 1098, 1105, 1084, 1072, 32, 1088, 1072, 1073, 1086, 1090) & ": " & JoinOperationNames(wsOut, 2, outRow - 1)
    wsOut.Cells(zRow + 3, 2).Value = "3. " & UW(1074, 1099, 1087, 1086, 1083, 1085, 1077, 1085, 1085, 1099, 1077, 32, 1088, 1072, 1073, 1086, 1090, 1099, 32, 1074, 32, 1088, 1072, 1084, 1082, 1072, 1093, 32, 1076, 1086, 1087, 1086, 1083, 1085, 1080, 1090, 1077, 1083, 1100, 1085, 1086, 1075, 1086, 32, 1086, 1073, 1098, 1105, 1084, 1072, 32, 1088, 1072, 1073, 1086, 1090) & ": " & zExtra
    wsOut.Cells(zRow + 4, 2).Value = "4. " & UW(1088, 1077, 1079, 1091, 1083, 1100, 1090, 1072, 1090, 1099, 32, 1080, 1089, 1087, 1099, 1090, 1072, 1085, 1080, 1081, 44, 32, 1090, 1077, 1089, 1090, 1086, 1074, 44, 32, 1079, 1072, 1084, 1077, 1088, 1086, 1074, 44, 32, 1080, 1085, 1089, 1087, 1077, 1082, 1094, 1080, 1081) & ": R" & UW(1080, 1079) & "=" & zRiz & " K=" & zK
    wsOut.Cells(zRow + 5, 2).Value = "5. " & UW(1086, 1090, 1082, 1083, 1086, 1085, 1077, 1085, 1080, 1103, 32, 1086, 1090, 32, 1058, 1050, 32, 1080, 32, 1088, 1077, 1082, 1086, 1084, 1077, 1085, 1076, 1072, 1094, 1080, 1080, 32, 1087, 1086, 32, 1082, 1086, 1088, 1088, 1077, 1082, 1090, 1080, 1088, 1086, 1074, 1082, 1077, 32, 1058, 1050) & ": " & zRec

    For zr = zRow To zRow + 5
        wsOut.Range(wsOut.Cells(zr, 2), wsOut.Cells(zr, 22)).Merge
        wsOut.Range(wsOut.Cells(zr, 2), wsOut.Cells(zr, 22)).Borders.LineStyle = xlContinuous
        wsOut.Range(wsOut.Cells(zr, 2), wsOut.Cells(zr, 22)).HorizontalAlignment = -4131
        wsOut.Rows(zr).RowHeight = 40
        wsOut.Range(wsOut.Cells(zr, 2), wsOut.Cells(zr, 22)).WrapText = True
    Next zr

    AppendResultToHistory wsHist, wsOut, wsIn, outRow - 1, zRow + 5, workerCount, lunch1, lunch2, lunchDurMin, primaryIsMin, clrEditHasColor, clrEditable, clrFont
    RefreshResultSheetColors wsOut, clrLocked, clrFont
    RefreshHistorySheetColors wsHist, clrLocked, clrEditHasColor, clrEditable, clrFont

    Dim chainEnd As Date, ceRow As Long
    chainEnd = 0
    For ceRow = 2 To outRow - 1
        Dim ce As Date
        ce = CDate(wsOut.Cells(ceRow, 20).Value) + wsOut.Cells(ceRow, 21).Value
        If ce > chainEnd Then chainEnd = ce
    Next ceRow
    wsIn.Range("B5").Value = DateValue(chainEnd)
    wsIn.Range("B6").Value = TimeValue(chainEnd)
    wsIn.Range("B5").NumberFormat = "dd"".""mm"".""yyyy"
    wsIn.Range("B6").NumberFormat = "hh:mm:ss"

    MsgBox UW(1043, 1086, 1090, 1086, 1074, 1086, 46, 32, 1057, 1092, 1086, 1088, 1084, 1080, 1088, 1086, 1074, 1072, 1085, 1086, 32, 1057, 1090, 1088, 1086, 1082, 58, 32) & (outRow - 2), vbInformation

    Dim errMsg As String
Cleanup:
    On Error Resume Next
    wsOut.Protect UW(49, 49, 52, 55, 48, 57)
    wsHist.Protect UW(49, 49, 52, 55, 48, 57)
    Application.EnableEvents = True
    wsIn.Cells(4, 13).Locked = False
    ApplyEditableStyle wsIn.Cells(4, 13), clrEditHasColor, clrEditable, clrFont
    wsIn.Protect UW(49, 49, 52, 55, 48, 57)
    On Error GoTo 0
    If Len(errMsg) > 0 Then
        MsgBox "Error: " & errMsg, vbCritical
    End If
    Exit Sub

EH:
    errMsg = Err.Description
    Resume Cleanup
End Sub

Private Sub EnsureSheetHeaders(ByVal wsOut As Worksheet, ByVal wsHist As Worksheet)
    wsOut.Cells(1, 1).Value = ""
    wsOut.Cells(1, 2).Value = UW(8470)
    wsOut.Cells(1, 3).Value = UW(1054, 1087, 1077, 1088, 1072, 1094, 1080, 1103)
    wsOut.Cells(1, 4).Value = UW(1054, 1073, 1077, 1076, 63)
    wsOut.Cells(1, 5).Value = UW(1055, 1072, 1091, 1079, 1072)
    wsOut.Cells(1, 6).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 32, 40, 1095, 1072, 1089, 41)
    wsOut.Cells(1, 7).Value = UW(1055, 1044, 1058, 1042)
    wsOut.Cells(1, 8).Value = "-"
    wsOut.Cells(1, 9).Value = "-"
    wsOut.Cells(1, 10).Value = "-"
    wsOut.Cells(1, 11).Value = "-"
    wsOut.Cells(1, 12).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 32, 40, 1084, 1080, 1085, 41)
    wsOut.Cells(1, 13).Value = "-"
    wsOut.Cells(1, 14).Value = "-"
    wsOut.Cells(1, 15).Value = UW(1044, 1072, 1090, 1072, 32, 1087, 1088, 1086, 1074, 1086, 1076, 1082, 1080)
    wsOut.Cells(1, 16).Value = UW(1048, 1089, 1087, 1086, 1083, 1085, 1080, 1090, 1077, 1083, 1100)
    wsOut.Cells(1, 17).Value = "-"
    wsOut.Cells(1, 18).Value = UW(1044, 1072, 1090, 1072, 32, 1053, 1072, 1095, 1072, 1083, 1072)
    wsOut.Cells(1, 19).Value = UW(1042, 1088, 1077, 1084, 1103, 32, 1053, 1072, 1095, 1072, 1083, 1072)
    wsOut.Cells(1, 20).Value = UW(1044, 1072, 1090, 1072, 32, 1050, 1086, 1085, 1094, 1072)
    wsOut.Cells(1, 21).Value = UW(1042, 1088, 1077, 1084, 1103, 32, 1050, 1086, 1085, 1094, 1072)
    wsOut.Cells(1, 22).Value = "INDEX"
    wsOut.Range("A1:V1").Font.Bold = True

    wsHist.Cells(1, 2).Value = UW(1048, 1089, 1090, 1086, 1088, 1080, 1103)
    wsHist.Cells(1, 2).Font.Bold = True
End Sub

Public Sub ClearResultAndHistory()
    If Not ZQ() Then
        MsgBox UW(1054, 1096, 1080, 1073, 1082, 1072, 33), vbCritical
        Exit Sub
    End If
    On Error GoTo EH
    Dim wsIn As Worksheet, wsR As Worksheet, wsH As Worksheet
    Set wsIn = ThisWorkbook.Worksheets(1)
    Set wsR = ThisWorkbook.Worksheets(2)
    Set wsH = ThisWorkbook.Worksheets(3)
    wsIn.Unprotect UW(49, 49, 52, 55, 48, 57)
    EnsureColorSettings wsIn
    wsIn.Range("B5").Value = Date
    wsIn.Range("B6").Value = TimeSerial(8, 0, 0)
    wsIn.Range("B7").Value = Date
    Dim clrLocked As Long, clrLockedFont As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrFont As Long
    ReadAllColors wsIn, clrLocked, clrLockedFont, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrFont
    wsIn.Protect UW(49, 49, 52, 55, 48, 57)
    wsR.Unprotect UW(49, 49, 52, 55, 48, 57)
    ClearResultArea wsR, clrLocked, clrFont
    wsR.Protect UW(49, 49, 52, 55, 48, 57)
    wsH.Unprotect UW(49, 49, 52, 55, 48, 57)
    ClearHistoryArea wsH, clrLocked, clrFont
    wsH.Protect UW(49, 49, 52, 55, 48, 57)
    RefreshWorkbookColors
    Exit Sub

Cleanup:
    On Error Resume Next
    wsIn.Protect UW(49, 49, 52, 55, 48, 57)
    wsR.Protect UW(49, 49, 52, 55, 48, 57)
    wsH.Protect UW(49, 49, 52, 55, 48, 57)
    Exit Sub
EH:
    Resume Cleanup
End Sub

Private Sub ClearResultArea(ByVal wsOut As Worksheet, ByVal clrLocked As Long, ByVal clrFont As Long)
    Dim rng As Range

    Set rng = BuildRowBoundRange(wsOut, 2, 2, 22, 2)
    On Error Resume Next
    rng.UnMerge
    On Error GoTo 0
    rng.ClearContents
    rng.Borders.LineStyle = xlNone
    rng.HorizontalAlignment = -4108
End Sub

Private Sub ClearHistoryArea(ByVal wsHist As Worksheet, ByVal clrLocked As Long, ByVal clrFont As Long)
    Dim rng As Range
    Dim lastRow As Long

    Set rng = BuildRowBoundRange(wsHist, 2, 2, 22, 2)
    lastRow = rng.Rows(rng.Rows.Count).Row
    Application.EnableEvents = False
    On Error Resume Next
    rng.UnMerge
    On Error GoTo 0
    rng.ClearContents
    rng.Borders.LineStyle = xlNone
    rng.HorizontalAlignment = -4108
    rng.VerticalAlignment = -4108
    rng.Font.Size = 14
    rng.Font.Color = clrFont
    rng.Font.Bold = False
    rng.NumberFormat = "General"
    rng.Locked = True
    rng.WrapText = False
    wsHist.Rows("2:" & lastRow).RowHeight = wsHist.StandardHeight
    Application.EnableEvents = True
End Sub

Public Sub SyncWorkerIdInputs(ByVal wsIn As Worksheet, ByVal workerCount As Long)
    Dim i As Long

    If workerCount < 1 Then workerCount = 1
    If workerCount > 10 Then workerCount = 10

    EnsureColorSettings wsIn
    Dim clrLocked As Long, clrLockedFont As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrFont As Long
    ReadAllColors wsIn, clrLocked, clrLockedFont, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrFont

    wsIn.Range("D3").Value = UW(1048, 1089, 1087, 1086, 1083, 1085, 1080, 1090, 1077, 1083, 1100)
    wsIn.Range("D3:E3").Font.Bold = True
    wsIn.Range("E4:E13").NumberFormat = "@"

    For i = 1 To 10
        Dim rowNum As Long
        rowNum = 3 + i

        wsIn.Cells(rowNum, 4).Value = UW(1048, 1089, 1087, 1086, 1083, 1085, 1080, 1090, 1077, 1083, 1100, 32) & i
        StyleWorkerInputRow wsIn, rowNum, (i <= workerCount), clrLocked, clrLockedFont, clrEditHasColor, clrEditable, clrFont
    Next i
End Sub

Public Sub SyncOperationRows(ByVal wsIn As Worksheet, ByVal opCount As Long)
    Application.EnableEvents = False
    On Error GoTo Cleanup

    If opCount < 1 Then opCount = 1
    If opCount > 20 Then opCount = 20

    EnsureColorSettings wsIn
    Dim clrLocked As Long, clrLockedFont As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrFont As Long
    ReadAllColors wsIn, clrLocked, clrLockedFont, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrFont

    Dim firstDurUnit As String, firstBreakUnit As String
    firstDurUnit = Trim$(CStr(wsIn.Cells(4, 12).Value))    ' L4
    firstBreakUnit = Trim$(CStr(wsIn.Cells(4, 14).Value))  ' N4

    Dim i As Long, r As Long, c As Long
    ' Editable operation columns: H(8), I(9), J(10), K(11), M(13), O(15)
    Dim editCols As Variant
    editCols = Array(8, 9, 10, 11, 13, 15)
    ' Locked synced columns: L(12), N(14)
    Dim syncCols As Variant
    syncCols = Array(12, 14)

    For i = 1 To opCount
        r = i + 3
        wsIn.Cells(r, 7).Value = i
        wsIn.Cells(r, 7).Font.Color = clrFont
        If Trim$(CStr(wsIn.Cells(r, 9).Value)) = "" Then
            wsIn.Cells(r, 9).Value = UW(1054, 1087, 1077, 1088, 1072, 1094, 1080, 1103) & " " & i
        End If
        If Trim$(CStr(wsIn.Cells(r, 11).Value)) = "" Then
            wsIn.Cells(r, 11).Value = 0
        End If
        If i > 1 Then
            If firstDurUnit <> "" Then wsIn.Cells(r, 12).Value = firstDurUnit
            wsIn.Cells(r, 13).Value = 0
            If firstBreakUnit <> "" Then wsIn.Cells(r, 14).Value = firstBreakUnit
            wsIn.Cells(r, 15).Value = ""
        End If
        ' Unlock editable columns
        For c = LBound(editCols) To UBound(editCols)
            wsIn.Cells(r, editCols(c)).Locked = False
            ApplyEditableStyle wsIn.Cells(r, editCols(c)), clrEditHasColor, clrEditable, clrFont
            wsIn.Cells(r, editCols(c)).Borders.LineStyle = xlContinuous
        Next c
        ' Lock pause (M) for first operation if no history exists
        If i = 1 Then
            Dim wsHist As Worksheet
            Set wsHist = ThisWorkbook.Worksheets(3)
            Dim histLastRow As Long
            histLastRow = wsHist.Cells(wsHist.Rows.Count, 2).End(xlUp).Row
            If histLastRow <= 1 Then
                wsIn.Cells(r, 13).Locked = True
                wsIn.Cells(r, 13).Interior.Color = clrLocked
            End If
        End If

        ' Synced columns: visible but locked (except row 2 = source)
        For c = LBound(syncCols) To UBound(syncCols)
            wsIn.Cells(r, syncCols(c)).Borders.LineStyle = xlContinuous
            wsIn.Cells(r, syncCols(c)).Font.Color = clrFont
            If i = 1 Then
                If clrEditHasColor Then
                    wsIn.Cells(r, syncCols(c)).Interior.Color = clrEditable
                Else
                    wsIn.Cells(r, syncCols(c)).Interior.Pattern = xlNone
                End If
                wsIn.Cells(r, syncCols(c)).Locked = False
            Else
                wsIn.Cells(r, syncCols(c)).Interior.Color = clrLocked
                wsIn.Cells(r, syncCols(c)).Locked = True
            End If
        Next c
    Next i

    ' Lock and color unused rows
    If opCount + 4 <= 23 Then
        Dim unusedRange As Range
        Set unusedRange = wsIn.Range(wsIn.Cells(opCount + 4, 7), wsIn.Cells(23, 15))
        unusedRange.ClearContents
        ApplyLockedStyle unusedRange, clrLocked, clrLockedFont
        unusedRange.Borders.LineStyle = xlNone
        unusedRange.Locked = True
    End If

Cleanup:
    Application.EnableEvents = True
End Sub

Public Sub SanitizeWorkerIdCell(ByVal targetCell As Range)
    If targetCell Is Nothing Then Exit Sub

    Dim normalized As String
    normalized = NormalizeWorkerIdText(CStr(targetCell.Value), 0, False)

    targetCell.NumberFormat = "00000000"
    If normalized = "" Then
        targetCell.ClearContents
    Else
        targetCell.Value = CLng(normalized)
    End If
End Sub

Private Sub StyleWorkerInputRow(ByVal wsIn As Worksheet, ByVal rowNum As Long, ByVal isVisible As Boolean, _
    ByVal clrLocked As Long, ByVal clrLockedFont As Long, _
    ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long, ByVal clrFont As Long)
    With wsIn.Cells(rowNum, 4)
        If isVisible Then
            .Font.Color = clrFont
        Else
            .Font.Color = clrLockedFont
            .ClearContents
        End If
        .Locked = True
    End With

    With wsIn.Cells(rowNum, 5)
        .NumberFormat = "00000000"
        If isVisible Then
            ApplyEditableStyle wsIn.Cells(rowNum, 5), clrEditHasColor, clrEditable, clrFont
            .Borders.LineStyle = xlContinuous
            .Locked = False
        Else
            ApplyLockedStyle wsIn.Cells(rowNum, 5), clrLocked, clrLockedFont
            .Borders.LineStyle = xlNone
            .Locked = True
            .ClearContents
        End If
    End With
End Sub

Private Function GetWorkerValue(ByVal wsIn As Worksheet, ByVal workerIndex As Long) As Variant
    Dim rawText As String
    rawText = Trim$(CStr(wsIn.Cells(3 + workerIndex, 5).Value))

    rawText = NormalizeWorkerIdText(rawText, workerIndex, True)
    If rawText = "" Then
        GetWorkerValue = workerIndex
    Else
        GetWorkerValue = CLng(rawText)
    End If
End Function

Private Function GetWorkerNumberFormat(ByVal wsIn As Worksheet, ByVal workerIndex As Long) As String
    Dim rawText As String
    rawText = Trim$(CStr(wsIn.Cells(3 + workerIndex, 5).Value))

    rawText = NormalizeWorkerIdText(rawText, 0, False)
    If rawText = "" Then
        GetWorkerNumberFormat = "General"
    Else
        GetWorkerNumberFormat = "00000000"
    End If
End Function

Private Function NormalizeWorkerIdText(ByVal rawText As String, ByVal workerIndex As Long, ByVal fallbackToIndex As Boolean) As String
    Dim digits As String
    digits = DigitsOnly(Trim$(rawText))

    If Len(digits) = 0 Then
        If fallbackToIndex Then
            NormalizeWorkerIdText = CStr(workerIndex)
        Else
            NormalizeWorkerIdText = ""
        End If
        Exit Function
    End If

    If Len(digits) > 8 Then digits = Left$(digits, 8)
    If Len(digits) < 8 Then digits = Right$(String$(8, "0") & digits, 8)

    NormalizeWorkerIdText = digits
End Function

Public Function NormalizeDecimal(ByVal rawText As String) As String
    Dim i As Long, ch As String
    Dim intPart As String, decPart As String
    Dim foundSep As Boolean

    foundSep = False
    intPart = ""
    decPart = ""

    For i = 1 To Len(rawText)
        ch = Mid$(rawText, i, 1)
        If ch >= "0" And ch <= "9" Then
            If foundSep Then
                decPart = decPart & ch
            Else
                intPart = intPart & ch
            End If
        ElseIf (ch = "," Or ch = ".") And Not foundSep Then
            foundSep = True
        End If
    Next i

    If Len(intPart) > 6 Then intPart = Left$(intPart, 6)
    If Len(decPart) > 2 Then decPart = Left$(decPart, 2)

    If intPart = "" And decPart = "" Then
        NormalizeDecimal = ""
    ElseIf decPart = "" Then
        NormalizeDecimal = intPart
    Else
        NormalizeDecimal = intPart & "," & decPart
    End If
End Function

Public Function NormalizeRizInput(ByVal rawText As String) As String
    Dim normalized As String
    normalized = NormalizeDecimal(rawText)
    If normalized = "" Then
        NormalizeRizInput = ""
    Else
        NormalizeRizInput = normalized & UW(32, 1052, 1086, 1084)
    End If
End Function

Public Function NormalizeKInput(ByVal rawText As String) As String
    NormalizeKInput = NormalizeDecimal(rawText)
End Function

Public Function DigitsOnly(ByVal rawText As String) As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(rawText)
        ch = Mid$(rawText, i, 1)
        If ch >= "0" And ch <= "9" Then
            DigitsOnly = DigitsOnly & ch
        End If
    Next i
End Function

Private Function TimeValueDefault(ByVal raw As Variant, ByVal Fallback As Date) As Date
    On Error GoTo Fallback
    If IsNumeric(raw) Then
        TimeValueDefault = CDbl(raw) - Fix(CDbl(raw))
    ElseIf IsDate(raw) Then
        TimeValueDefault = TimeValue(raw)
    Else
        TimeValueDefault = TimeValue(CStr(raw))
    End If
    Exit Function
Fallback:
    TimeValueDefault = Fallback
End Function

Private Function ConvertDurationToDays(ByVal valueNum As Double, ByVal unitName As String) As Double
    If unitName = "hour" Then
        ConvertDurationToDays = valueNum / 24#
    Else
        ConvertDurationToDays = valueNum / 1440#
    End If
End Function

Private Function NormalizeUnit(ByVal rawUnit As Variant) As String
    Dim unitName As String
    unitName = LCase$(Trim$(CStr(rawUnit)))
    If unitName = "hour" Or unitName = LCase$(UW(1095, 1072, 1089)) Or unitName = LCase$(UW(1095)) Then
        NormalizeUnit = "hour"
    Else
        NormalizeUnit = "min"
    End If
End Function

Private Sub ParseLunchParams( _
    ByVal lunch1Raw As Variant, _
    ByVal lunch2Raw As Variant, _
    ByRef lunchDurMin As Double, _
    ByRef lunch1 As Date, _
    ByRef lunch2 As Date, _
    ByRef hasLunch2 As Boolean, _
    ByRef lunchDurDays As Double)

    lunch1 = TimeValueDefault(lunch1Raw, TimeSerial(12, 0, 0))
    lunch2 = TimeValueDefault(lunch2Raw, TimeSerial(0, 0, 0))
    hasLunch2 = (Format$(lunch2, "hh:nn:ss") <> "00:00:00")
    If lunchDurMin < 0 Then lunchDurMin = 0
    If lunchDurMin > 480 Then lunchDurMin = 480
    lunchDurDays = lunchDurMin / 1440#
End Sub

Private Function ComputeEndWithLunch( _
    ByVal opStart As Date, _
    ByVal durationDays As Double, _
    ByVal lunch1 As Date, _
    ByVal lunch2 As Date, _
    ByVal hasLunch2 As Boolean, _
    ByVal lunchDurDays As Double, _
    ByRef crossedLunch As Boolean) As Date

    Dim opEnd As Date
    opEnd = opStart + durationDays

    Dim lStart As Date, lEnd As Date
    Dim daySerial As Long

    ' 1. Check Lunch 1
    For daySerial = CLng(Int(opStart)) - 1 To CLng(Int(opEnd)) + 1
        lStart = daySerial + TimePart(lunch1)
        lEnd = lStart + lunchDurDays

        If opStart < lStart And opEnd > lStart Then
            opEnd = opEnd + lunchDurDays
            crossedLunch = True
            Exit For
        End If
    Next daySerial

    ' 2. Check Lunch 2
    If hasLunch2 Then
        For daySerial = CLng(Int(opStart)) - 1 To CLng(Int(opEnd)) + 1
            lStart = daySerial + TimePart(lunch2)
            lEnd = lStart + lunchDurDays

            If opStart < lStart And opEnd > lStart Then
                opEnd = opEnd + lunchDurDays
                crossedLunch = True
                Exit For
            End If
        Next daySerial
    End If

    ComputeEndWithLunch = opEnd
End Function

Private Function ShiftStartOutOfLunch( _
    ByVal startDateTime As Date, _
    ByVal lunch1 As Date, _
    ByVal lunch2 As Date, _
    ByVal hasLunch2 As Boolean, _
    ByVal lunchDurDays As Double) As Date

    Dim st As Date
    st = startDateTime

    Dim daySerial As Long
    ' Check lunch 1
    For daySerial = CLng(Int(st)) - 1 To CLng(Int(st)) + 1
        Dim lStart As Date, lEnd As Date
        lStart = daySerial + TimePart(lunch1)
        lEnd = lStart + lunchDurDays
        If st >= lStart - 0.00001 And st < lEnd Then
            st = lEnd
            Exit For
        End If
    Next daySerial

    ' Check lunch 2
    If hasLunch2 Then
        For daySerial = CLng(Int(st)) - 1 To CLng(Int(st)) + 1
            Dim l2Start As Date, l2End As Date
            l2Start = daySerial + TimePart(lunch2)
            l2End = l2Start + lunchDurDays
            If st >= l2Start - 0.00001 And st < l2End Then
                st = l2End
                Exit For
            End If
        Next daySerial
    End If

    ShiftStartOutOfLunch = st
End Function

' ==========================================
' PURE EXCEL FORMULA BUILDERS
' ==========================================
Private Function BuildLunchShiftFormulaStr(ByVal rawTimeExpr As String, _
    ByVal lunch1 As Date, ByVal lunch2 As Date, ByVal lunchDurMin As Double) As String

    Dim lh As Long, lm As Long
    Dim lh2 As Long, lm2 As Long
    Dim ld As Long
    lh = Hour(lunch1): lm = Minute(lunch1)
    lh2 = Hour(lunch2): lm2 = Minute(lunch2)
    ld = CLng(lunchDurMin)

    Dim l1Val As String, l1End As String
    l1Val = "TIME(" & lh & "," & lm & ",0)"
    l1End = "(TIME(" & lh & "," & lm & ",0)+TIME(0," & ld & ",0))"

    Dim tp As String, cond1 As String, res1 As String, shifted1 As String
        tp = "ROUND(MOD(" & rawTimeExpr & ", 1), 6)"
        Dim l1vR As String, l1eR As String
        l1vR = "ROUND(" & l1Val & ", 6)"
        l1eR = "ROUND(" & l1End & ", 6)"

        cond1 = "AND(" & tp & ">=" & l1vR & ", " & tp & "<" & l1eR & ")"
    res1 = "(INT(" & rawTimeExpr & ") + " & l1End & ")"
    shifted1 = "IF(" & cond1 & "," & res1 & "," & rawTimeExpr & ")"

    Dim hasLunch2 As Boolean
    hasLunch2 = (Format$(lunch2, "hh:nn:ss") <> "00:00:00")

    If hasLunch2 Then
        Dim l2Val As String, l2End As String
        Dim tp2 As String, cond2 As String, res2 As String
        l2Val = "TIME(" & lh2 & "," & lm2 & ",0)"
        l2End = "(TIME(" & lh2 & "," & lm2 & ",0)+TIME(0," & ld & ",0))"

            tp2 = "ROUND(MOD(" & shifted1 & ", 1), 6)"
            Dim l2vR As String, l2eR As String
            l2vR = "ROUND(" & l2Val & ", 6)"
            l2eR = "ROUND(" & l2End & ", 6)"
            
            cond2 = "AND(" & tp2 & ">=" & l2vR & ", " & tp2 & "<" & l2eR & ")"
        res2 = "(INT(" & shifted1 & ") + " & l2End & ")"

        BuildLunchShiftFormulaStr = "IF(" & cond2 & "," & res2 & "," & shifted1 & ")"
    Else
        BuildLunchShiftFormulaStr = shifted1
    End If
End Function

Private Sub BuildEndFormulasStr(ByVal startDateExpr As String, ByVal startTimeExpr As String, _
    ByVal durExpr As String, ByVal unitDiv As String, _
    ByVal lunch1 As Date, ByVal lunch2 As Date, ByVal lunchDurMin As Double, _
    ByRef outEndDateFormula As String, ByRef outEndTimeFormula As String)

    Dim lh As Long, lm As Long
    Dim lh2 As Long, lm2 As Long
    Dim ld As Long
    lh = Hour(lunch1): lm = Minute(lunch1)
    lh2 = Hour(lunch2): lm2 = Minute(lunch2)
    ld = CLng(lunchDurMin)

    Dim l1Val As String, l1End As String, lDurVal As String
    l1Val = "TIME(" & lh & "," & lm & ",0)"
    l1End = "(TIME(" & lh & "," & lm & ",0)+TIME(0," & ld & ",0))"
    lDurVal = "TIME(0," & ld & ",0)"

    Dim hasLunch2 As Boolean
    hasLunch2 = (Format$(lunch2, "hh:nn:ss") <> "00:00:00")

    Dim stMod As String, rawEndRel As String, enC1 As String, enShift1 As String, mainMath As String
        stMod = "ROUND(MOD(" & startTimeExpr & ", 1), 6)"
        rawEndRel = "ROUND(" & stMod & "+(" & durExpr & "/" & unitDiv & "), 6)"

        Dim l1vR As String, l1vCR As String, l1v1R As String, l1v1CR As String
        l1vR = "ROUND(" & l1Val & ", 6)"
        l1vCR = "ROUND(" & l1Val & "+TIME(0,0,1), 6)"
        l1v1R = "ROUND((" & l1Val & "+1), 6)"
        l1v1CR = "ROUND((" & l1Val & "+1)+TIME(0,0,1), 6)"

        enC1 = "OR(AND(" & stMod & "<" & l1vR & ", " & rawEndRel & ">=" & l1vCR & "), AND(" & stMod & "<" & l1v1R & ", " & rawEndRel & ">=" & l1v1CR & "))"
    enShift1 = "IF(" & enC1 & ", " & lDurVal & ", 0)"
    mainMath = "(" & startDateExpr & ") + (" & startTimeExpr & ") + (" & durExpr & "/" & unitDiv & ") + " & enShift1

    If hasLunch2 Then
        Dim l2Val As String, l2End As String
        l2Val = "TIME(" & lh2 & "," & lm2 & ",0)"
        l2End = "(TIME(" & lh2 & "," & lm2 & ",0)+TIME(0," & ld & ",0))"

        Dim stMod2 As String, rawEndRel2 As String, enC2 As String, enShift2 As String
            stMod2 = "ROUND(MOD((" & startTimeExpr & ") + " & enShift1 & ", 1), 6)"
            rawEndRel2 = "ROUND(" & stMod2 & "+(" & durExpr & "/" & unitDiv & "), 6)"

            Dim l2vR As String, l2vCR As String, l2v1R As String, l2v1CR As String
            l2vR = "ROUND(" & l2Val & ", 6)"
            l2vCR = "ROUND(" & l2Val & "+TIME(0,0,1), 6)"
            l2v1R = "ROUND((" & l2Val & "+1), 6)"
            l2v1CR = "ROUND((" & l2Val & "+1)+TIME(0,0,1), 6)"

            enC2 = "OR(AND(" & stMod2 & "<" & l2vR & ", " & rawEndRel2 & ">=" & l2vCR & "), AND(" & stMod2 & "<" & l2v1R & ", " & rawEndRel2 & ">=" & l2v1CR & "))"
        enShift2 = "IF(" & enC2 & ", " & lDurVal & ", 0)"

        outEndDateFormula = "INT(" & mainMath & " + " & enShift2 & ")"
        outEndTimeFormula = "MOD(" & mainMath & " + " & enShift2 & ", 1)"
    Else
        outEndDateFormula = "INT(" & mainMath & ")"
        outEndTimeFormula = "MOD(" & mainMath & ", 1)"
    End If
End Sub

Private Function BuildLunchIconFormulaStr(ByVal startTimeExpr As String, ByVal durExpr As String, ByVal unitDiv As String, _
    ByVal lunch1 As Date, ByVal lunch2 As Date, ByVal lunchDurMin As Double, ByVal daStr As String) As String

    Dim lh As Long, lm As Long
    Dim lh2 As Long, lm2 As Long
    Dim ld As Long
    lh = Hour(lunch1): lm = Minute(lunch1)
    lh2 = Hour(lunch2): lm2 = Minute(lunch2)
    ld = CLng(lunchDurMin)

    Dim l1Val As String, l1End As String, lDurVal As String
    l1Val = "TIME(" & lh & "," & lm & ",0)"
    l1End = "(TIME(" & lh & "," & lm & ",0)+TIME(0," & ld & ",0))"
    lDurVal = "TIME(0," & ld & ",0)"

    Dim hasLunch2 As Boolean
    hasLunch2 = (Format$(lunch2, "hh:nn:ss") <> "00:00:00")

    Dim startTimeMod As String, endTimeRel As String, l1EndMod As String
    Dim icWasShifted1 As String, icCovers1 As String, icC1 As String, icShift1 As String

        startTimeMod = "ROUND(MOD(" & startTimeExpr & ", 1), 6)"
        endTimeRel = "ROUND(" & startTimeMod & "+(" & durExpr & "/" & unitDiv & "), 6)"
        l1EndMod = "ROUND(MOD(" & l1End & ", 1), 6)"

        Dim l1vR As String, l1vCR As String, l1v1R As String, l1v1CR As String
        l1vR = "ROUND(" & l1Val & ", 6)"
        l1vCR = "ROUND(" & l1Val & "+TIME(0,0,1), 6)"
        l1v1R = "ROUND((" & l1Val & "+1), 6)"
        l1v1CR = "ROUND((" & l1Val & "+1)+TIME(0,0,1), 6)"

        icWasShifted1 = "ROUND(ABS(" & startTimeMod & "-" & l1EndMod & "), 6)<ROUND(TIME(0,0,1), 6)"
        icCovers1 = "OR(AND(" & startTimeMod & "<" & l1vR & ", " & endTimeRel & ">" & l1vCR & "), AND(" & startTimeMod & "<" & l1v1R & ", " & endTimeRel & ">" & l1v1CR & "))"
    icC1 = "OR(" & icWasShifted1 & ", " & icCovers1 & ")"
    icShift1 = "IF(" & icC1 & ", " & lDurVal & ", 0)"

    If hasLunch2 Then
        Dim l2Val As String, l2End As String
        l2Val = "TIME(" & lh2 & "," & lm2 & ",0)"
        l2End = "(TIME(" & lh2 & "," & lm2 & ",0)+TIME(0," & ld & ",0))"

        Dim shiftedStartMod As String, shiftedEndRel As String, l2EndMod As String
        Dim icWasShifted2 As String, icCovers2 As String, icC2 As String

            shiftedStartMod = "ROUND(MOD(" & startTimeExpr & "+" & icShift1 & ", 1), 6)"
            shiftedEndRel = "ROUND(" & shiftedStartMod & "+(" & durExpr & "/" & unitDiv & "), 6)"
            l2EndMod = "ROUND(MOD(" & l2End & ", 1), 6)"

            Dim l2vR As String, l2vCR As String, l2v1R As String, l2v1CR As String
            l2vR = "ROUND(" & l2Val & ", 6)"
            l2vCR = "ROUND(" & l2Val & "+TIME(0,0,1), 6)"
            l2v1R = "ROUND((" & l2Val & "+1), 6)"
            l2v1CR = "ROUND((" & l2Val & "+1)+TIME(0,0,1), 6)"

            icWasShifted2 = "ROUND(ABS(" & shiftedStartMod & "-" & l2EndMod & "), 6)<ROUND(TIME(0,0,1), 6)"
            icCovers2 = "OR(AND(" & shiftedStartMod & "<" & l2vR & ", " & shiftedEndRel & ">" & l2vCR & "), AND(" & shiftedStartMod & "<" & l2v1R & ", " & shiftedEndRel & ">" & l2v1CR & "))"
        icC2 = "OR(" & icWasShifted2 & ", " & icCovers2 & ")"

        BuildLunchIconFormulaStr = "IF(OR(" & icC1 & ", " & icC2 & "), """ & daStr & """, """")"
    Else
        BuildLunchIconFormulaStr = "IF(" & icC1 & ", """ & daStr & """, """")"
    End If
End Function

Private Function TimePart(ByVal dateTimeValue As Date) As Double
    TimePart = CDbl(dateTimeValue) - Fix(CDbl(dateTimeValue))
End Function

Private Function IsWorkerSelected(ByVal spec As String, ByVal workerIndex As Long) As Boolean
    spec = Replace(spec, " ", "")
    If spec = "" Then
        IsWorkerSelected = True
        Exit Function
    End If

    Dim parts() As String
    parts = Split(spec, ",")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim token As String
        token = Trim$(parts(i))
        If token = "" Then GoTo ContinueToken

        If InStr(1, token, "-", vbTextCompare) > 0 Then
            Dim r() As String
            r = Split(token, "-")
            If UBound(r) = 1 Then
                Dim fromN As Long, toN As Long
                fromN = CLng(val(r(0)))
                toN = CLng(val(r(1)))
                If fromN > toN Then
                    Dim tmp As Long
                    tmp = fromN
                    fromN = toN
                    toN = tmp
                End If
                If workerIndex >= fromN And workerIndex <= toN Then
                    IsWorkerSelected = True
                    Exit Function
                End If
            End If
        Else
            If workerIndex = CLng(val(token)) Then
                IsWorkerSelected = True
                Exit Function
            End If
        End If
ContinueToken:
    Next i

    IsWorkerSelected = False
End Function

Private Function JoinOperationNames(ByVal wsOut As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long) As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = firstRow To lastRow
        Dim opName As String
        opName = Trim$(CStr(wsOut.Cells(r, 3).Value))
        If opName <> "" Then
            If Not dict.Exists(opName) Then dict.Add opName, 1
        End If
    Next r

    If dict.Count = 0 Then
        JoinOperationNames = ""
    Else
        JoinOperationNames = Join(dict.Keys, ", ")
    End If
End Function

Private Sub AppendResultToHistory( _
    ByVal wsHist As Worksheet, _
    ByVal wsOut As Worksheet, _
    ByVal wsIn As Worksheet, _
    ByVal lastDataRow As Long, _
    ByVal lastZ7Row As Long, _
    ByVal workerCount As Long, _
    ByVal lunch1 As Date, _
    ByVal lunch2 As Date, _
    ByVal lunchDurMin As Double, _
    ByVal primaryIsMin As Boolean, _
    ByVal clrEditHasColor As Boolean, _
    ByVal clrEditable As Long, _
    ByVal clrFont As Long)

    Dim NextRow As Long
    NextRow = wsHist.Cells(wsHist.Rows.Count, 2).End(xlUp).Row + 1
    If NextRow < 2 Then NextRow = 2

    ' пустая строка-разделитель между блоками
    NextRow = NextRow + 1

    ' заголовок блока: Номер Заказа | Наименование | Сформирован
    Dim orderNum As String, orderName As String
    orderNum = Trim$(CStr(wsIn.Range("B3").Value))
    orderName = Trim$(CStr(wsIn.Range("B4").Value))
    wsHist.Range(wsHist.Cells(NextRow, 2), wsHist.Cells(NextRow, 22)).Merge
    wsHist.Cells(NextRow, 2).Value = UW(1053, 1086, 1084, 1077, 1088, 32, 1047, 1072, 1082, 1072, 1079, 1072) & ": " & orderNum & _
        "  |  " & UW(1053, 1072, 1080, 1084, 1077, 1085, 1086, 1074, 1072, 1085, 1080, 1077) & ": " & orderName & _
        "  |  " & UW(1057, 1092, 1086, 1088, 1084, 1080, 1088, 1086, 1074, 1072, 1085) & ": " & Format$(Now, "dd.mm.yyyy hh:nn:ss")
    wsHist.Cells(NextRow, 2).Font.Bold = True
    wsHist.Range(wsHist.Cells(NextRow, 2), wsHist.Cells(NextRow, 22)).Borders.LineStyle = xlContinuous
    wsHist.Range(wsHist.Cells(NextRow, 2), wsHist.Cells(NextRow, 22)).HorizontalAlignment = -4108
    NextRow = NextRow + 1

    wsOut.Range("A1:V1").Copy Destination:=wsHist.Cells(NextRow, 1)
    NextRow = NextRow + 1
    Dim dataStartRow As Long, dataEndRow As Long
    dataStartRow = NextRow
    dataEndRow = NextRow + (lastDataRow - 2)

    wsOut.Range("A2:V" & lastDataRow).Copy Destination:=wsHist.Cells(dataStartRow, 1)

    ' Фиксируем якорь даты проводки внутри блока History (snapshot на момент копирования):
    wsHist.Cells(dataStartRow, 15).Value = wsOut.Cells(2, 15).Value
    If dataEndRow > dataStartRow Then
        wsHist.Range("O" & (dataStartRow + 1) & ":O" & dataEndRow).Formula = "=$O$" & dataStartRow
    End If

    ' первая строка нового блока должна начинаться от конца последней операции предыдущего блока.
    Dim prevDataEndRow As Long, scanRow As Long
    prevDataEndRow = 0
    For scanRow = dataStartRow - 1 To 2 Step -1
        If IsNumeric(wsHist.Cells(scanRow, 2).Value) Then
            If Len(Trim$(CStr(wsHist.Cells(scanRow, 20).Value))) > 0 And Len(Trim$(CStr(wsHist.Cells(scanRow, 21).Value))) > 0 Then
                prevDataEndRow = scanRow
                Exit For
            End If
        End If
    Next scanRow

    If prevDataEndRow > 0 Then
        Dim rawDTFirst As String
        rawDTFirst = "(T" & prevDataEndRow & "+U" & prevDataEndRow & "+E" & dataStartRow & ")"
        Dim shiftFFirst As String
        shiftFFirst = BuildLunchShiftFormulaStr(rawDTFirst, lunch1, lunch2, lunchDurMin)
        wsHist.Cells(dataStartRow, 18).Formula = "=INT(" & shiftFFirst & ")"
        wsHist.Cells(dataStartRow, 19).Formula = "=MOD(" & shiftFFirst & ",1)"
    Else
        wsHist.Cells(dataStartRow, 18).Value = wsOut.Cells(2, 18).Value
        wsHist.Cells(dataStartRow, 19).Value = wsOut.Cells(2, 19).Value
    End If

    wsHist.Range("D" & dataStartRow & ":D" & dataEndRow).NumberFormat = "General"

    Dim lDivisor As String
    If primaryIsMin Then
        lDivisor = "1440"
    Else
        lDivisor = "24"
    End If

    Dim rr As Long
    For rr = dataStartRow To dataEndRow
        If rr > dataStartRow Then
            Dim rawDT As String
            rawDT = "(T" & (rr - 1) & "+U" & (rr - 1) & "+E" & rr & ")"
            Dim shiftF As String
            shiftF = BuildLunchShiftFormulaStr(rawDT, lunch1, lunch2, lunchDurMin)
            wsHist.Cells(rr, 18).Formula = "=IF(B" & rr & "=B" & (rr - 1) & ",R" & (rr - 1) & ",INT(" & shiftF & "))"
            wsHist.Cells(rr, 19).Formula = "=IF(B" & rr & "=B" & (rr - 1) & ",S" & (rr - 1) & ",MOD(" & shiftF & ",1))"
        End If

        Dim endD As String, endT As String
        BuildEndFormulasStr "R" & rr, "S" & rr, "L" & rr, lDivisor, lunch1, lunch2, lunchDurMin, endD, endT
        wsHist.Cells(rr, 20).Formula = "=" & endD
        wsHist.Cells(rr, 21).Formula = "=" & endT
    Next rr

    For rr = dataStartRow To dataEndRow
        Dim iconF As String
        iconF = BuildLunchIconFormulaStr("S" & rr, "L" & rr, lDivisor, lunch1, lunch2, lunchDurMin, UW(1044, 1040))
        wsHist.Cells(rr, 4).Formula = "=" & iconF
    Next rr

    wsHist.Range("E" & dataStartRow & ":E" & dataEndRow).NumberFormat = "h:mm:ss"
    wsHist.Range("O" & dataStartRow & ":O" & dataEndRow).NumberFormat = "dd"".""mm"".""yyyy"
    wsHist.Range("R" & dataStartRow & ":R" & dataEndRow).NumberFormat = "dd"".""mm"".""yyyy"
    wsHist.Range("S" & dataStartRow & ":S" & dataEndRow).NumberFormat = "h:mm:ss"
    wsHist.Range("T" & dataStartRow & ":T" & dataEndRow).NumberFormat = "dd"".""mm"".""yyyy"
    wsHist.Range("U" & dataStartRow & ":U" & dataEndRow).NumberFormat = "h:mm:ss"

    ' Borders for column header row and data rows
    Dim headerRow As Long
    headerRow = dataStartRow - 1
    wsHist.Range(wsHist.Cells(headerRow, 2), wsHist.Cells(dataEndRow, 22)).Borders.LineStyle = xlContinuous

    ' Thick borders on left of G and right of U for the data block (header + data)
    wsHist.Range(wsHist.Cells(headerRow, 7), wsHist.Cells(dataEndRow, 7)).Borders(7).LineStyle = xlContinuous
    wsHist.Range(wsHist.Cells(headerRow, 7), wsHist.Cells(dataEndRow, 7)).Borders(7).Weight = 4
    wsHist.Range(wsHist.Cells(headerRow, 21), wsHist.Cells(dataEndRow, 21)).Borders(10).LineStyle = xlContinuous
    wsHist.Range(wsHist.Cells(headerRow, 21), wsHist.Cells(dataEndRow, 21)).Borders(10).Weight = 4

    NextRow = NextRow + (lastDataRow - 1) + 1

    ' Z7 block with merged cells
    Dim z7Start As Long
    z7Start = NextRow
    wsHist.Cells(NextRow, 2).Value = "Z7"
    wsHist.Cells(NextRow + 1, 2).Resize(5, 1).Value = wsOut.Range("B" & (lastZ7Row - 4) & ":B" & lastZ7Row).Value

    Dim zhr As Long
    For zhr = z7Start To z7Start + 5
        wsHist.Range(wsHist.Cells(zhr, 2), wsHist.Cells(zhr, 22)).Merge
        wsHist.Range(wsHist.Cells(zhr, 2), wsHist.Cells(zhr, 22)).Borders.LineStyle = xlContinuous
        wsHist.Range(wsHist.Cells(zhr, 2), wsHist.Cells(zhr, 22)).HorizontalAlignment = -4131
        wsHist.Rows(zhr).RowHeight = 40
        wsHist.Range(wsHist.Cells(zhr, 2), wsHist.Cells(zhr, 22)).WrapText = True
    Next zhr

    ' Unlock editable cells (non-formula) in columns E, G, L, P
    Dim editCol As Variant, ec As Long
    editCol = Array(5, 7, 12, 16)
    For ec = LBound(editCol) To UBound(editCol)
        For rr = dataStartRow To dataEndRow
            ' Skip Pause (col 5) for first row of first block
            If editCol(ec) = 5 And rr = dataStartRow And prevDataEndRow = 0 Then GoTo SkipCell
            If Not wsHist.Cells(rr, editCol(ec)).HasFormula Then
                wsHist.Cells(rr, editCol(ec)).Locked = False
                If clrEditHasColor Then
                    wsHist.Cells(rr, editCol(ec)).Interior.Color = clrEditable
                Else
                    wsHist.Cells(rr, editCol(ec)).Interior.Pattern = xlNone
                End If
                If editCol(ec) = 12 Then wsHist.Cells(rr, editCol(ec)).NumberFormat = "@"
            End If
SkipCell:
        Next rr
    Next ec

    ' Apply text format to all L column cells for consistent display
    wsHist.Range(wsHist.Cells(dataStartRow, 12), wsHist.Cells(dataEndRow, 12)).NumberFormat = "@"

    ' R+S first row: unlock only if value (no previous block)
    If prevDataEndRow = 0 Then
        wsHist.Cells(dataStartRow, 18).Locked = False
        If clrEditHasColor Then
            wsHist.Cells(dataStartRow, 18).Interior.Color = clrEditable
        Else
            wsHist.Cells(dataStartRow, 18).Interior.Pattern = xlNone
        End If
        wsHist.Cells(dataStartRow, 19).Locked = False
        If clrEditHasColor Then
            wsHist.Cells(dataStartRow, 19).Interior.Color = clrEditable
        Else
            wsHist.Cells(dataStartRow, 19).Interior.Pattern = xlNone
        End If
    End If

    ' Font 14 for all new data (copied from Result may have different font)
    Dim blockStart As Long
    blockStart = headerRow - 1  ' title row
    wsHist.Range(wsHist.Cells(blockStart, 1), wsHist.Cells(z7Start + 5, 22)).Font.Size = 14

    ' Fixed column widths for font 14
    wsHist.Columns("A:A").ColumnWidth = 2.7
    wsHist.Columns("B:B").ColumnWidth = 3.5
    wsHist.Columns("C:C").ColumnWidth = 60
    wsHist.Columns("D:D").ColumnWidth = 8
    wsHist.Columns("E:E").ColumnWidth = 11
    wsHist.Columns("F:F").ColumnWidth = 15
    wsHist.Columns("G:G").ColumnWidth = 12
    wsHist.Columns("H:H").ColumnWidth = 1.5
    wsHist.Columns("I:I").ColumnWidth = 1.5
    wsHist.Columns("J:J").ColumnWidth = 1.5
    wsHist.Columns("K:K").ColumnWidth = 1.5
    wsHist.Columns("L:L").ColumnWidth = 16
    wsHist.Columns("M:M").ColumnWidth = 1.5
    wsHist.Columns("N:N").ColumnWidth = 1.5
    wsHist.Columns("O:O").ColumnWidth = 18
    wsHist.Columns("P:P").ColumnWidth = 16
    wsHist.Columns("Q:Q").ColumnWidth = 1.5
    wsHist.Columns("R:R").ColumnWidth = 16
    wsHist.Columns("S:S").ColumnWidth = 16
    wsHist.Columns("T:T").ColumnWidth = 16
    wsHist.Columns("U:U").ColumnWidth = 16
    wsHist.Columns("V:V").ColumnWidth = 10
End Sub

Private Function IsNumericOrder(ByVal s As String) As Boolean
    Dim ch As Long
    If Len(s) = 0 Then IsNumericOrder = False: Exit Function
    For ch = 1 To Len(s)
        If Mid(s, ch, 1) < "0" Or Mid(s, ch, 1) > "9" Then
            IsNumericOrder = False: Exit Function
        End If
    Next ch
    IsNumericOrder = True
End Function

Public Sub LoadMRS()
    If Not ZQ() Then
        MsgBox UW(1054, 1096, 1080, 1073, 1082, 1072, 33), vbCritical
        Exit Sub
    End If
    On Error GoTo EH

    Dim stage As String
    stage = "LoadMRS start"

    Dim wsMRS As Worksheet
    Set wsMRS = ThisWorkbook.Worksheets(4)

    Dim filePath As Variant
    If MsgBox(UW(1042, 1099, 1075, 1088, 1091, 1078, 1072, 1081, 1090, 1077, 32, 1076, 1072, 1085, 1085, 1099, 1077, 32, 1079, 1072, 32, 1054, 1044, 1048, 1053, 32, 1076, 1077, 1085, 1100, 44, 32, 1080, 1085, 1072, 1095, 1077, 32, 1073, 1091, 1076, 1077, 1090, 32, 1073, 1077, 1076, 1072, 46, 46, 46), vbOKCancel + vbExclamation) = vbCancel Then Exit Sub

    filePath = Application.GetOpenFilename(UW(1060, 1072, 1081, 1083, 1099) & " Excel (*.xlsx), *.xlsx")
    If filePath = False Then Exit Sub

    Dim wsIn As Worksheet
    Set wsIn = ThisWorkbook.Worksheets(1)
    stage = "Ensure color settings"
    wsIn.Unprotect UW(49, 49, 52, 55, 48, 57)
    EnsureColorSettings wsIn
    Dim clrLocked As Long, clrLockedFont As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrFont As Long
    stage = "Read colors from settings"
    ReadAllColors wsIn, clrLocked, clrLockedFont, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrFont

    Dim lunch1 As Date, lunch2 As Date, hasLunch2 As Boolean, lunchDurDays As Double
    Dim lunchDurMin As Double
    lunchDurMin = CDbl(Val(wsIn.Range("B12").Value))
    ParseLunchParams wsIn.Range("B10").Value, wsIn.Range("B11").Value, lunchDurMin, lunch1, lunch2, hasLunch2, lunchDurDays

    wsIn.Protect UW(49, 49, 52, 55, 48, 57)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    wsMRS.Unprotect UW(49, 49, 52, 55, 48, 57)
    stage = "Clear MRS sheet"
    ClearMRSArea wsMRS, clrLocked, clrFont

    Dim srcWb As Workbook
    stage = "Open source workbook"
    Set srcWb = Workbooks.Open(CStr(filePath), ReadOnly:=True)
    Dim srcWs As Worksheet
    Set srcWs = srcWb.Sheets(1)

    Dim lastRow As Long
    stage = "Detect source last row"
    lastRow = srcWs.Cells(srcWs.Rows.Count, 5).End(xlUp).Row
    If lastRow < 2 Then
        srcWb.Close False
        MsgBox UW(1060, 1072, 1081, 1083, 32, 1087, 1091, 1089, 1090), vbInformation
        GoTo Done
    End If

    Dim data As Variant
    stage = "Read source range into array"
    data = srcWs.Range("A1:Z" & lastRow).Value
    srcWb.Close False

    Dim totalRows As Long
    totalRows = UBound(data, 1) - 1

    ' Parse rows into arrays
    stage = "Parse source rows"
    Dim i As Long
    Dim arrOrder() As String, arrOp() As String
    Dim arrWName() As String, arrWID() As String
    Dim arrSDate() As Date, arrSTime() As Date
    Dim arrEDate() As Date, arrETime() As Date
    Dim arrDur() As Double
    Dim arrWPlace() As String, arrPDTV() As String

    ReDim arrOrder(1 To totalRows)
    ReDim arrOp(1 To totalRows)
    ReDim arrWName(1 To totalRows)
    ReDim arrWID(1 To totalRows)
    ReDim arrSDate(1 To totalRows)
    ReDim arrSTime(1 To totalRows)
    ReDim arrEDate(1 To totalRows)
    ReDim arrETime(1 To totalRows)
    ReDim arrDur(1 To totalRows)
    ReDim arrWPlace(1 To totalRows)
    ReDim arrPDTV(1 To totalRows)

    Dim assignVal As String, slashPos As Long
    Dim validCount As Long
    Dim tmpOrder As String
    validCount = 0
    For i = 1 To totalRows
        assignVal = CStr(data(i + 1, 5))
        slashPos = InStr(assignVal, "/")
        If slashPos > 0 Then
            tmpOrder = Left(assignVal, slashPos - 1)
        Else
            tmpOrder = assignVal
        End If
        If Not IsNumericOrder(tmpOrder) Then GoTo NextRow

        validCount = validCount + 1
        If slashPos > 0 Then
            arrOrder(validCount) = tmpOrder
            arrOp(validCount) = Mid(assignVal, slashPos + 1)
        Else
            arrOrder(validCount) = tmpOrder
            arrOp(validCount) = ""
        End If
        arrWName(validCount) = CStr(data(i + 1, 3))
        arrWID(validCount) = CStr(data(i + 1, 4))
        If Not IsEmpty(data(i + 1, 8)) Then
            If IsDate(data(i + 1, 8)) Then
                arrSDate(validCount) = CDate(data(i + 1, 8))
            ElseIf IsNumeric(data(i + 1, 8)) Then
                arrSDate(validCount) = CDate(CDbl(data(i + 1, 8)))
            End If
        End If
        If Not IsEmpty(data(i + 1, 9)) Then
            If IsDate(data(i + 1, 9)) Then
                arrSTime(validCount) = CDate(data(i + 1, 9))
            ElseIf IsNumeric(data(i + 1, 9)) Then
                arrSTime(validCount) = CDate(CDbl(data(i + 1, 9)))
            End If
        End If
        If Not IsEmpty(data(i + 1, 10)) Then
            If IsDate(data(i + 1, 10)) Then
                arrEDate(validCount) = CDate(data(i + 1, 10))
            ElseIf IsNumeric(data(i + 1, 10)) Then
                arrEDate(validCount) = CDate(CDbl(data(i + 1, 10)))
            End If
        End If
        If Not IsEmpty(data(i + 1, 11)) Then
            If IsDate(data(i + 1, 11)) Then
                arrETime(validCount) = CDate(data(i + 1, 11))
            ElseIf IsNumeric(data(i + 1, 11)) Then
                arrETime(validCount) = CDate(CDbl(data(i + 1, 11)))
            End If
        End If
        If Not IsEmpty(data(i + 1, 12)) Then
            If IsNumeric(data(i + 1, 12)) Then arrDur(validCount) = CDbl(data(i + 1, 12))
        End If
        arrWPlace(validCount) = CStr(data(i + 1, 14))
        arrPDTV(validCount) = CStr(data(i + 1, 16))
NextRow:
    Next i

    totalRows = validCount
    If totalRows = 0 Then
        MsgBox UW(1053, 1077, 1090, 32, 1079, 1072, 1082, 1072, 1079, 1086, 1074, 32, 1089, 32, 1095, 1080, 1089, 1083, 1086, 1074, 1099, 1084, 32, 1085, 1086, 1084, 1077, 1088, 1086, 1084), vbInformation
        GoTo Done
    End If

    ' === Step 0: Collect unique dates ===
    stage = "Collect unique dates"
    Dim dictDates As Object
    Set dictDates = CreateObject("Scripting.Dictionary")
    For i = 1 To totalRows
        Dim dateKey As Long
        dateKey = CLng(arrSDate(i))
        If Not dictDates.Exists(dateKey) Then
            dictDates.Add dateKey, arrSDate(i)
        End If
    Next i

    Dim dateCount As Long
    dateCount = dictDates.Count
    Dim dateKeys() As Long
    ReDim dateKeys(1 To dateCount)
    Dim dk As Variant
    Dim kd As Long: kd = 0
    For Each dk In dictDates.Keys
        kd = kd + 1
        dateKeys(kd) = CLng(dk)
    Next dk
    For i = 1 To dateCount - 1
        Dim jd As Long
        For jd = 1 To dateCount - i
            If dateKeys(jd) > dateKeys(jd + 1) Then
                Dim tmpDateKey As Long
                tmpDateKey = dateKeys(jd): dateKeys(jd) = dateKeys(jd + 1): dateKeys(jd + 1) = tmpDateKey
            End If
        Next jd
    Next i

    Dim outRow As Long
    outRow = 2

    Dim d As Long
    For d = 1 To dateCount
        Dim curDateKey As Long
        curDateKey = dateKeys(d)

        ' Date header
        stage = "Write date header " & CStr(d)
        outRow = outRow + 1
        wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Merge
        wsMRS.Cells(outRow, 2).Value = Format(CDate(CDbl(curDateKey)), "dd"".""mm"".""yyyy")
        SetCellBoldSafe wsMRS.Cells(outRow, 2), True, "MRS date header"
        wsMRS.Cells(outRow, 2).Font.Size = 18
        wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
        wsMRS.Cells(outRow, 2).HorizontalAlignment = -4108
        wsMRS.Cells(outRow, 2).Interior.Color = clrMrsHeader
        wsMRS.Rows(outRow).RowHeight = 40

        ' === Step 1: Build worker set per order (filtered by date) ===
        stage = "Build order-worker map for date " & CStr(d)
        Dim dictOrderW As Object
        Set dictOrderW = CreateObject("Scripting.Dictionary")
        Dim dictOrderTime As Object
        Set dictOrderTime = CreateObject("Scripting.Dictionary")

        For i = 1 To totalRows
            If CLng(arrSDate(i)) <> curDateKey Then GoTo SkipDateRow
            Dim startDbl As Double
        startDbl = CDbl(arrSDate(i)) + CDbl(arrSTime(i))
        If Not dictOrderW.Exists(arrOrder(i)) Then
            Dim dw As Object
            Set dw = CreateObject("Scripting.Dictionary")
            dictOrderW.Add arrOrder(i), dw
            dictOrderTime.Add arrOrder(i), startDbl
        Else
            If startDbl < dictOrderTime(arrOrder(i)) Then
                dictOrderTime(arrOrder(i)) = startDbl
            End If
        End If
        If Not dictOrderW(arrOrder(i)).Exists(arrWID(i)) Then
            dictOrderW(arrOrder(i)).Add arrWID(i), arrWName(i)
        End If
SkipDateRow:
        Next i

    ' === Step 2: Build brigade key per order (sorted worker IDs) ===
    stage = "Build brigade keys for date " & CStr(d)
    Dim dictOrdBrigKey As Object ' order -> brigade key string
    Set dictOrdBrigKey = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    For Each key In dictOrderW.Keys
        Dim wIds As Variant
        wIds = dictOrderW(key).Keys
        Dim wCount As Long
        wCount = dictOrderW(key).Count
        ' Sort worker IDs (bubble sort on Variant array)
        Dim j As Long, tmpS As String
        If wCount > 1 Then
            For i = 0 To wCount - 2
                For j = 0 To wCount - 2 - i
                    If CStr(wIds(j)) > CStr(wIds(j + 1)) Then
                        tmpS = wIds(j): wIds(j) = wIds(j + 1): wIds(j + 1) = tmpS
                    End If
                Next j
            Next i
        End If
        Dim bKey As String: bKey = ""
        For i = 0 To wCount - 1
            If i > 0 Then bKey = bKey & ","
            bKey = bKey & CStr(wIds(i))
        Next i
        dictOrdBrigKey.Add CStr(key), bKey
    Next key

    ' === Step 3: Group orders by brigade key ===
    stage = "Group brigades for date " & CStr(d)
    Dim dictBrigades As Object ' brigKey -> Dictionary of order -> startTime
    Set dictBrigades = CreateObject("Scripting.Dictionary")
    Dim dictBrigTime As Object ' brigKey -> earliest start time across all orders
    Set dictBrigTime = CreateObject("Scripting.Dictionary")

    For Each key In dictOrdBrigKey.Keys
        Dim bk As String
        bk = dictOrdBrigKey(key)
        If Not dictBrigades.Exists(bk) Then
            Dim db As Object
            Set db = CreateObject("Scripting.Dictionary")
            dictBrigades.Add bk, db
            dictBrigTime.Add bk, dictOrderTime(key)
        Else
            If dictOrderTime(key) < dictBrigTime(bk) Then
                dictBrigTime(bk) = dictOrderTime(key)
            End If
        End If
        dictBrigades(bk).Add CStr(key), CDbl(dictOrderTime(key))
    Next key

    ' === Step 4: Sort brigades by earliest start time ===
    stage = "Sort brigades for date " & CStr(d)
    Dim brigCount As Long
    brigCount = dictBrigades.Count
    Dim brigKeys() As String, brigTimes() As Double
    ReDim brigKeys(1 To brigCount)
    ReDim brigTimes(1 To brigCount)

    Dim k As Long: k = 0
    For Each key In dictBrigades.Keys
        k = k + 1
        brigKeys(k) = CStr(key)
        brigTimes(k) = CDbl(dictBrigTime(key))
    Next key

    Dim tmpD As Double
    For i = 1 To brigCount - 1
        For j = 1 To brigCount - i
            If brigTimes(j) > brigTimes(j + 1) Then
                tmpD = brigTimes(j): brigTimes(j) = brigTimes(j + 1): brigTimes(j + 1) = tmpD
                tmpS = brigKeys(j): brigKeys(j) = brigKeys(j + 1): brigKeys(j + 1) = tmpS
            End If
        Next j
    Next i

    ' === Step 5: Write blocks ===
    stage = "Write brigade blocks for date " & CStr(d)
    Dim b As Long
    For b = 1 To brigCount
        Dim curBrigKey As String
        curBrigKey = brigKeys(b)

        ' Build worker names string from first order in this brigade
        Dim brigWorkerIds As Variant
        brigWorkerIds = Split(curBrigKey, ",")
        Dim firstOrdKey As Variant
        For Each firstOrdKey In dictBrigades(curBrigKey).Keys: Exit For: Next
        Dim wStr As String
        wStr = ""
        Dim wi As Long
        For wi = 0 To UBound(brigWorkerIds)
            If wi > 0 Then wStr = wStr & ", "
            Dim wName As String: wName = ""
            If dictOrderW(CStr(firstOrdKey)).Exists(CStr(brigWorkerIds(wi))) Then
                wName = dictOrderW(CStr(firstOrdKey))(CStr(brigWorkerIds(wi)))
            End If
            wStr = wStr & wName & " (" & brigWorkerIds(wi) & ")"
        Next wi

        ' Brigade header (merged B:N)
        stage = "Write brigade header " & CStr(d) & "." & CStr(b)
        outRow = outRow + 1
        wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Merge
        wsMRS.Cells(outRow, 2).Value = UW(1041, 1088, 1080, 1075, 1072, 1076, 1072) & ": " & wStr
        SetCellBoldSafe wsMRS.Cells(outRow, 2), True, "MRS brigade header"
        wsMRS.Cells(outRow, 2).Font.Size = 16
        wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
        wsMRS.Cells(outRow, 2).HorizontalAlignment = -4108  ' xlCenter
        wsMRS.Cells(outRow, 2).Interior.Color = clrMrsSub
        wsMRS.Cells(outRow, 2).WrapText = True
        wsMRS.Rows(outRow).RowHeight = 60

        ' Sort orders within brigade by start time
        Dim ordInBrig As Long
        ordInBrig = dictBrigades(curBrigKey).Count
        Dim ordKeys() As String, ordTimes() As Double
        ReDim ordKeys(1 To ordInBrig)
        ReDim ordTimes(1 To ordInBrig)
        k = 0
        Dim oKey As Variant
        For Each oKey In dictBrigades(curBrigKey).Keys
            k = k + 1
            ordKeys(k) = CStr(oKey)
            ordTimes(k) = CDbl(dictBrigades(curBrigKey)(oKey))
        Next oKey
        For i = 1 To ordInBrig - 1
            For j = 1 To ordInBrig - i
                If ordTimes(j) > ordTimes(j + 1) Then
                    tmpD = ordTimes(j): ordTimes(j) = ordTimes(j + 1): ordTimes(j + 1) = tmpD
                    tmpS = ordKeys(j): ordKeys(j) = ordKeys(j + 1): ordKeys(j + 1) = tmpS
                End If
            Next j
        Next i

        ' Per-worker tracking: workerID -> last output row number
        Dim dictWorkerLastRow As Object
        Set dictWorkerLastRow = CreateObject("Scripting.Dictionary")

        ' For each order in brigade
        Dim oi As Long
        Dim pauseCellRow As Long

        For oi = 1 To ordInBrig
            Dim curOrder As String
            curOrder = ordKeys(oi)

            ' Collect row indices for this order
            Dim blkIdx() As Long, blkCnt As Long
            blkCnt = 0
            For i = 1 To totalRows
                If arrOrder(i) = curOrder And CLng(arrSDate(i)) = curDateKey Then
                    blkCnt = blkCnt + 1
                    ReDim Preserve blkIdx(1 To blkCnt)
                    blkIdx(blkCnt) = i
                End If
            Next i

            ' Sort by start time, then operation, then worker ID
            Dim doSwap As Boolean
            Dim tmpL As Long
            Dim tA As Double, tB As Double
            For i = 1 To blkCnt - 1
                For j = 1 To blkCnt - i
                    doSwap = False
                    tA = CDbl(arrSDate(blkIdx(j))) + CDbl(arrSTime(blkIdx(j)))
                    tB = CDbl(arrSDate(blkIdx(j + 1))) + CDbl(arrSTime(blkIdx(j + 1)))
                    If tA > tB Then
                        doSwap = True
                    ElseIf tA = tB Then
                        If arrOp(blkIdx(j)) > arrOp(blkIdx(j + 1)) Then
                            doSwap = True
                        ElseIf arrOp(blkIdx(j)) = arrOp(blkIdx(j + 1)) Then
                            If arrWID(blkIdx(j)) > arrWID(blkIdx(j + 1)) Then doSwap = True
                        End If
                    End If
                    If doSwap Then
                        tmpL = blkIdx(j): blkIdx(j) = blkIdx(j + 1): blkIdx(j + 1) = tmpL
                    End If
                Next j
            Next i

            Dim fIdx As Long
            fIdx = blkIdx(1)

            ' Pause row between orders (not before first)
            If oi > 1 Then
                stage = "Write pause row " & CStr(d) & "." & CStr(b) & "." & CStr(oi)
                outRow = outRow + 1
                pauseCellRow = outRow
                wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 8)).Merge
                wsMRS.Cells(outRow, 2).Value = UW(1055, 1072, 1091, 1079, 1072, 32, 40, 1084, 1080, 1085, 41, 58)
                SetCellBoldSafe wsMRS.Cells(outRow, 2), True, "MRS pause row"
                wsMRS.Cells(outRow, 2).HorizontalAlignment = -4108  ' xlCenter
                wsMRS.Cells(outRow, 9).NumberFormat = "0.00"
                wsMRS.Cells(outRow, 9).Value = 0
                wsMRS.Cells(outRow, 9).NumberFormat = "@"
                wsMRS.Cells(outRow, 9).Locked = False
                If clrEditHasColor Then
                    wsMRS.Cells(outRow, 9).Interior.Color = clrEditable
                Else
                    wsMRS.Cells(outRow, 9).Interior.Pattern = xlNone
                End If
                wsMRS.Range(wsMRS.Cells(outRow, 10), wsMRS.Cells(outRow, 14)).Merge
                wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
            End If

            ' Order header (merged B:N)
            stage = "Write order header " & CStr(d) & "." & CStr(b) & "." & CStr(oi)
            outRow = outRow + 1
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Merge
            wsMRS.Cells(outRow, 2).Value = UW(1047, 1072, 1082, 1072, 1079) & ": " & curOrder & " | " & UW(1056, 1072, 1073, 1086, 1095, 1077, 1077, 32, 1084, 1077, 1089, 1090, 1086) & ": " & arrWPlace(fIdx)
            wsMRS.Cells(outRow, 2).Font.Size = 15
            SetCellBoldSafe wsMRS.Cells(outRow, 2), True, "MRS order header"
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
            wsMRS.Cells(outRow, 2).HorizontalAlignment = -4108  ' xlCenter

            ' Column headers row 1
            stage = "Write column header block " & CStr(d) & "." & CStr(b) & "." & CStr(oi)
            outRow = outRow + 1
            Dim hdrRow1 As Long: hdrRow1 = outRow
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow + 1, 2)).Merge
            wsMRS.Cells(outRow, 2).Value = UW(8470)
            wsMRS.Range(wsMRS.Cells(outRow, 3), wsMRS.Cells(outRow + 1, 3)).Merge
            wsMRS.Cells(outRow, 3).Value = UW(1054, 1087, 1077, 1088, 1072, 1094, 1080, 1103)
            wsMRS.Range(wsMRS.Cells(outRow, 4), wsMRS.Cells(outRow + 1, 4)).Merge
            wsMRS.Cells(outRow, 4).Value = UW(1048, 1089, 1087, 1086, 1083, 1085, 1080, 1090, 1077, 1083, 1100)
            wsMRS.Range(wsMRS.Cells(outRow, 5), wsMRS.Cells(outRow + 1, 5)).Merge
            wsMRS.Cells(outRow, 5).Value = UW(1058, 1072, 1073, 46, 8470)
            wsMRS.Range(wsMRS.Cells(outRow, 6), wsMRS.Cells(outRow + 1, 6)).Merge
            wsMRS.Cells(outRow, 6).Value = UW(1044, 1072, 1090, 1072, 32, 1085, 1072, 1095, 1072, 1083, 1072)
            wsMRS.Range(wsMRS.Cells(outRow, 7), wsMRS.Cells(outRow + 1, 7)).Merge
            wsMRS.Cells(outRow, 7).Value = UW(1042, 1088, 1077, 1084, 1103, 32, 1085, 1072, 1095, 1072, 1083, 1072)
            wsMRS.Range(wsMRS.Cells(outRow, 8), wsMRS.Cells(outRow + 1, 8)).Merge
            wsMRS.Cells(outRow, 8).Value = UW(1044, 1072, 1090, 1072, 32, 1082, 1086, 1085, 1094, 1072)
            wsMRS.Range(wsMRS.Cells(outRow, 9), wsMRS.Cells(outRow + 1, 9)).Merge
            wsMRS.Cells(outRow, 9).Value = UW(1042, 1088, 1077, 1084, 1103, 32, 1082, 1086, 1085, 1094, 1072)
            wsMRS.Range(wsMRS.Cells(outRow, 10), wsMRS.Cells(outRow, 11)).Merge
            wsMRS.Cells(outRow, 10).Value = UW(1055, 1072, 1088, 1089, 1080, 1085, 1075, 40, 1044, 1074, 1086, 1081, 1085, 1086, 1077, 32, 1054, 1082, 1088, 1091, 1075, 1083, 1077, 1085, 1080, 1077, 41)
            wsMRS.Range(wsMRS.Cells(outRow, 12), wsMRS.Cells(outRow, 13)).Merge
            wsMRS.Cells(outRow, 12).Value = UW(1056, 1072, 1089, 1095, 1077, 1090, 1085, 1086, 1077)
            wsMRS.Range(wsMRS.Cells(outRow, 14), wsMRS.Cells(outRow + 1, 14)).Merge
            wsMRS.Cells(outRow, 14).Value = UW(1054, 1073, 1077, 1076, 63)

            ' Column headers row 2
            outRow = outRow + 1
            wsMRS.Cells(outRow, 10).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 40, 1095, 1072, 1089, 41)
            wsMRS.Cells(outRow, 11).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 40, 1084, 1080, 1085, 41)
            wsMRS.Cells(outRow, 12).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 40, 1095, 1072, 1089, 41)
            wsMRS.Cells(outRow, 13).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 40, 1084, 1080, 1085, 41)

            ' Header formatting
            SetRangeBoldSafe wsMRS.Range(wsMRS.Cells(hdrRow1, 2), wsMRS.Cells(outRow, 14)), True, "MRS column header block"
            wsMRS.Range(wsMRS.Cells(hdrRow1, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
            wsMRS.Range(wsMRS.Cells(hdrRow1, 2), wsMRS.Cells(outRow, 14)).WrapText = True

            ' Track workers seen in current order (for pause logic)
            Dim dictWorkerSeenInOrder As Object
            Set dictWorkerSeenInOrder = CreateObject("Scripting.Dictionary")

            ' Data rows
            Dim totalDur As Double: totalDur = 0
            Dim dataStartRow As Long: dataStartRow = outRow + 1
            Dim idx As Long
            Dim daStr As String: daStr = UW(1044, 1040)
            For i = 1 To blkCnt
                stage = "Write data row " & CStr(d) & "." & CStr(b) & "." & CStr(oi) & "." & CStr(i)
                outRow = outRow + 1
                idx = blkIdx(i)
                Dim curWID As String
                curWID = arrWID(idx)

                wsMRS.Cells(outRow, 2).Value = i
                wsMRS.Cells(outRow, 3).Value = arrOp(idx)
                wsMRS.Cells(outRow, 4).Value = arrWName(idx)
                wsMRS.Cells(outRow, 5).Value = arrWID(idx)

                ' Start date/time: per-worker chain (shift via ShiftStartOutOfLunch UDF)
                Dim rawDT As String
                If dictWorkerLastRow.Exists(curWID) Then
                    Dim prevR As Long
                    prevR = CLng(dictWorkerLastRow(curWID))
                    If Not dictWorkerSeenInOrder.Exists(curWID) Then
                        ' Cross-order: previous end + pause
                        rawDT = "(H" & prevR & "+I" & prevR & "+I" & pauseCellRow & "/1440)"
                    Else
                        ' Within order: start = previous end for same worker
                        rawDT = "(H" & prevR & "+I" & prevR & ")"
                    End If
                    Dim shiftF As String
                    shiftF = BuildLunchShiftFormulaStr(rawDT, lunch1, lunch2, lunchDurMin)
                    wsMRS.Cells(outRow, 6).Formula = "=INT(" & shiftF & ")"
                    wsMRS.Cells(outRow, 7).Formula = "=MOD(" & shiftF & ",1)"
                Else
                    ' First time seeing this worker: parsed values (G editable, shift if in lunch)
                    wsMRS.Cells(outRow, 6).Value = arrSDate(idx)
                    Dim startTimeVal As Date
                    startTimeVal = arrSTime(idx)
                    Dim stFull As Date
                    stFull = arrSDate(idx) + startTimeVal
                    Dim stShifted As Date
                    stShifted = ShiftStartOutOfLunch(stFull, lunch1, lunch2, hasLunch2, lunchDurDays)
                    wsMRS.Cells(outRow, 7).Value = stShifted - Int(stShifted)
                    wsMRS.Cells(outRow, 7).Locked = False
                    If clrEditHasColor Then
                        wsMRS.Cells(outRow, 7).Interior.Color = clrEditable
                    Else
                        wsMRS.Cells(outRow, 7).Interior.Pattern = xlNone
                    End If
                End If
                wsMRS.Cells(outRow, 6).NumberFormat = "dd"".""mm"".""yyyy"
                wsMRS.Cells(outRow, 7).NumberFormat = "h:mm:ss"

                ' Parsing duration (static from source)
                wsMRS.Cells(outRow, 10).Value = arrDur(idx)
                wsMRS.Cells(outRow, 11).Value = arrDur(idx) * 60

                ' M = Расчетное мин (INPUT, editable) — initial value from parsed times
                Dim initMin As Double
                initMin = (CDbl(arrEDate(idx)) + CDbl(arrETime(idx)) - CDbl(arrSDate(idx)) - CDbl(arrSTime(idx))) * 1440
                If arrSDate(idx) = arrEDate(idx) Then
                    If arrSTime(idx) <= lunch1 And arrETime(idx) > lunch1 Then
                        initMin = initMin - lunchDurMin
                    End If
                    If hasLunch2 And arrSTime(idx) <= lunch2 And arrETime(idx) > lunch2 Then
                        initMin = initMin - lunchDurMin
                    End If
                End If
                If initMin < 0 Then initMin = 0
                wsMRS.Cells(outRow, 13).NumberFormat = "0.00"
                wsMRS.Cells(outRow, 13).Value = Round(initMin, 2)
                wsMRS.Cells(outRow, 13).NumberFormat = "@"
                wsMRS.Cells(outRow, 13).Locked = False
                If clrEditHasColor Then
                    wsMRS.Cells(outRow, 13).Interior.Color = clrEditable
                Else
                    wsMRS.Cells(outRow, 13).Interior.Pattern = xlNone
                End If

                ' L = Расчетное час (formula from M)
                Dim rStr As String: rStr = CStr(outRow)
                wsMRS.Cells(outRow, 12).Formula = "=M" & rStr & "/60"
                wsMRS.Cells(outRow, 12).NumberFormat = "0.00"

                ' H = End date (formula: start + duration + lunch)
                ' I = End time (formula: start + duration + lunch)
                Dim endD As String, endT As String
                BuildEndFormulasStr "F" & rStr, "G" & rStr, "M" & rStr, "1440", lunch1, lunch2, lunchDurMin, endD, endT
                wsMRS.Cells(outRow, 8).Formula = "=" & endD
                wsMRS.Cells(outRow, 8).NumberFormat = "dd"".""mm"".""yyyy"
                wsMRS.Cells(outRow, 9).Formula = "=" & endT
                wsMRS.Cells(outRow, 9).NumberFormat = "h:mm:ss"

                ' N = Обед? (formula)
                Dim iconF As String
                iconF = BuildLunchIconFormulaStr("G" & rStr, "M" & rStr, "1440", lunch1, lunch2, lunchDurMin, daStr)
                wsMRS.Cells(outRow, 14).Formula = "=" & iconF

                wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
                totalDur = totalDur + arrDur(idx)

                ' Update worker tracking
                dictWorkerSeenInOrder(curWID) = True
                If dictWorkerLastRow.Exists(curWID) Then
                    dictWorkerLastRow(curWID) = outRow
                Else
                    dictWorkerLastRow.Add curWID, outRow
                End If
            Next i

            ' Totals row with SUM formulas
            stage = "Write totals row " & CStr(d) & "." & CStr(b) & "." & CStr(oi)
            Dim dataEndRow As Long: dataEndRow = outRow
            outRow = outRow + 1
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 9)).Merge
            wsMRS.Cells(outRow, 2).Value = UW(1048, 1090, 1086, 1075, 1086) & ": " & blkCnt & " " & UW(1079, 1072, 1087, 1080, 1089, 1077, 1081)
            SetCellBoldSafe wsMRS.Cells(outRow, 2), True, "MRS totals label"
            wsMRS.Cells(outRow, 2).HorizontalAlignment = -4108  ' xlCenter
            wsMRS.Cells(outRow, 10).Formula = "=SUM(J" & dataStartRow & ":J" & dataEndRow & ")"
            wsMRS.Cells(outRow, 10).NumberFormat = "0.00"
            SetCellBoldSafe wsMRS.Cells(outRow, 10), True, "MRS totals J"
            wsMRS.Cells(outRow, 11).Formula = "=SUM(K" & dataStartRow & ":K" & dataEndRow & ")"
            wsMRS.Cells(outRow, 11).NumberFormat = "0.00"
            SetCellBoldSafe wsMRS.Cells(outRow, 11), True, "MRS totals K"
            wsMRS.Cells(outRow, 12).Formula = "=SUM(L" & dataStartRow & ":L" & dataEndRow & ")"
            wsMRS.Cells(outRow, 12).NumberFormat = "0.00"
            SetCellBoldSafe wsMRS.Cells(outRow, 12), True, "MRS totals L"
            wsMRS.Cells(outRow, 13).Formula = "=SUM(M" & dataStartRow & ":M" & dataEndRow & ")"
            wsMRS.Cells(outRow, 13).NumberFormat = "0.00"
            SetCellBoldSafe wsMRS.Cells(outRow, 13), True, "MRS totals M"
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
        Next oi

        outRow = outRow + 1 ' empty row between brigades
    Next b

    Next d

    ' Apply formatting to used area
    stage = "Finalize MRS formatting"
    Dim dataRange As Range
    Set dataRange = wsMRS.Range(wsMRS.Cells(3, 2), wsMRS.Cells(outRow, 14))
    dataRange.Font.Size = 14
    dataRange.Font.Color = clrFont
    dataRange.HorizontalAlignment = -4108
    dataRange.VerticalAlignment = -4108

    ' Restore header font sizes and alignment (overwritten by bulk formatting above)
    Dim rr As Long
    Dim cellVal As String
    Dim mergeColCount As Long
    For rr = 3 To outRow
        If wsMRS.Cells(rr, 2).MergeCells Then
            mergeColCount = wsMRS.Cells(rr, 2).MergeArea.Columns.Count
            cellVal = CStr(wsMRS.Cells(rr, 2).Value)
            If mergeColCount >= 10 Then
                If Left$(cellVal, 1) >= "0" And Left$(cellVal, 1) <= "9" Then
                    ' Date header (e.g. "02.02.2026")
                    wsMRS.Cells(rr, 2).Font.Size = 18
                ElseIf Left$(cellVal, 1) = ChrW(1041) Then
                    ' Brigade header (starts with "Б" from "Бригада:")
                    wsMRS.Cells(rr, 2).Font.Size = 16
                ElseIf Left$(cellVal, 1) = ChrW(1047) Then
                    ' Order header (starts with "З" from "Заказ:")
                    wsMRS.Cells(rr, 2).Font.Size = 15
                End If
            End If
        End If
    Next rr

    RefreshMRSSheetColors wsMRS, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrFont

    ' Column widths
    wsMRS.Columns("A:A").ColumnWidth = 2.7
    wsMRS.Columns("B:B").ColumnWidth = 4
    wsMRS.Columns("C:C").ColumnWidth = 13
    wsMRS.Columns("D:D").ColumnWidth = 45
    wsMRS.Columns("E:E").ColumnWidth = 12
    wsMRS.Columns("F:F").ColumnWidth = 15
    wsMRS.Columns("G:G").ColumnWidth = 17
    wsMRS.Columns("H:H").ColumnWidth = 15
    wsMRS.Columns("I:I").ColumnWidth = 17
    wsMRS.Columns("J:J").ColumnWidth = 24
    wsMRS.Columns("K:K").ColumnWidth = 24
    wsMRS.Columns("L:L").ColumnWidth = 24
    wsMRS.Columns("M:M").ColumnWidth = 24
    wsMRS.Columns("N:N").ColumnWidth = 10

    MsgBox UW(1055, 1072, 1088, 1089, 1080, 1085, 1075, 32, 1079, 1072, 1074, 1077, 1088, 1096, 1077, 1085, 32, 1091, 1089, 1087, 1077, 1096, 1085, 1086, 46), vbInformation

    Dim errMsg As String
Done:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close False
    wsMRS.Protect UW(49, 49, 52, 55, 48, 57)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    If Len(errMsg) > 0 Then
        MsgBox UW(1054, 1096, 1080, 1073, 1082, 1072, 58, 32) & errMsg, vbCritical
    End If
    Exit Sub
EH:
    errMsg = stage & " | " & Err.Description
    Resume Done
End Sub

Public Sub ClearMRS()
    If Not ZQ() Then
        MsgBox UW(1054, 1096, 1080, 1073, 1082, 1072, 33), vbCritical
        Exit Sub
    End If
    On Error GoTo EH
    Dim wsIn As Worksheet, wsMRS As Worksheet
    Set wsIn = ThisWorkbook.Worksheets(1)
    Set wsMRS = ThisWorkbook.Worksheets(4)
    wsIn.Unprotect UW(49, 49, 52, 55, 48, 57)
    EnsureColorSettings wsIn
    Dim clrLocked As Long, clrLockedFont As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrFont As Long
    ReadAllColors wsIn, clrLocked, clrLockedFont, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrFont
    wsIn.Protect UW(49, 49, 52, 55, 48, 57)
    wsMRS.Unprotect UW(49, 49, 52, 55, 48, 57)
    ClearMRSArea wsMRS, clrLocked, clrFont
    wsMRS.Protect UW(49, 49, 52, 55, 48, 57)
    RefreshWorkbookColors
    Exit Sub

Cleanup:
    On Error Resume Next
    wsIn.Protect UW(49, 49, 52, 55, 48, 57)
    wsMRS.Protect UW(49, 49, 52, 55, 48, 57)
    Exit Sub
EH:
    Resume Cleanup
End Sub

Private Sub ClearMRSArea(ByVal wsMRS As Worksheet, ByVal clrLocked As Long, ByVal clrFont As Long)
    Dim rng As Range
    Dim lastRow As Long

    Set rng = BuildRowBoundRange(wsMRS, 3, 2, 14, 3)
    lastRow = rng.Rows(rng.Rows.Count).Row
    On Error Resume Next
    rng.UnMerge
    On Error GoTo 0
    rng.ClearContents
    rng.Borders.LineStyle = xlNone
    rng.HorizontalAlignment = -4108
    rng.VerticalAlignment = -4108
    rng.Font.Size = 14
    rng.Font.Color = clrFont
    rng.Font.Bold = False
    rng.NumberFormat = "General"
    rng.Locked = True
    rng.WrapText = False
    wsMRS.Rows("3:" & lastRow).RowHeight = wsMRS.StandardHeight
End Sub

'@

$inputSheetCode = @'
Option Explicit

Private mColorSignature As String

Private Sub Worksheet_Activate()
    mColorSignature = BuildColorSettingsSignature(Me)
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo SafeExit

    Dim currentSignature As String
    currentSignature = BuildColorSettingsSignature(Me)

    If mColorSignature = "" Then
        mColorSignature = currentSignature
    ElseIf StrComp(currentSignature, mColorSignature, vbBinaryCompare) <> 0 Then
        Application.EnableEvents = False
        RefreshWorkbookColors
        mColorSignature = BuildColorSettingsSignature(Me)
    End If

SafeExit:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    If Target Is Nothing Then Exit Sub

    Application.EnableEvents = False
    Me.Unprotect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)

    If Not Intersect(Target, Me.Range("B8")) Is Nothing Then
        Dim opCount As Long
        opCount = CLng(Val(Me.Range("B8").Value))
        If opCount < 1 Then opCount = 1
        If opCount > 20 Then opCount = 20
        Me.Range("B8").Value = opCount
        SyncOperationRows Me, opCount
    End If

    If Not Intersect(Target, Me.Range("B9")) Is Nothing Then
        Dim workerCount As Long
        workerCount = CLng(Val(Me.Range("B9").Value))
        If workerCount < 1 Then workerCount = 1
        If workerCount > 10 Then workerCount = 10
        Me.Range("B9").Value = workerCount
        SyncWorkerIdInputs Me, workerCount
    End If

    Dim workerRange As Range, cell As Range
    Set workerRange = Intersect(Target, Me.Range("E4:E13"))
    If Not workerRange Is Nothing Then
        For Each cell In workerRange.Cells
            SanitizeWorkerIdCell cell
        Next cell
    End If

    If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
        Dim b1digits As String
        b1digits = DigitsOnly(CStr(Me.Range("B3").Value))
        If Len(b1digits) > 12 Then b1digits = Left$(b1digits, 12)
        Me.Range("B3").NumberFormat = "000000000000"
        If b1digits = "" Then
            Me.Range("B3").ClearContents
        Else
            Me.Range("B3").Value = CDbl(b1digits)
        End If
    End If

    Dim hRange As Range, hCell As Range
    Set hRange = Intersect(Target, Me.Range("H4:H23"))
    If Not hRange Is Nothing Then
        For Each hCell In hRange.Cells
            Dim hDigits As String
            hDigits = DigitsOnly(CStr(hCell.Value))
            If Len(hDigits) > 8 Then hDigits = Left$(hDigits, 8)
            hCell.NumberFormat = "00000000"
            If hDigits = "" Then
                hCell.ClearContents
            Else
                hCell.Value = CLng(hDigits)
            End If
        Next hCell
    End If

    Dim jRange As Range, jCell As Range
    Set jRange = Intersect(Target, Me.Range("J4:J23"))
    If Not jRange Is Nothing Then
        For Each jCell In jRange.Cells
            Dim rawJ As String
            rawJ = CStr(jCell.Value)
            If InStr(rawJ, ".") > 0 Then
                rawJ = Replace(rawJ, ".", ",")
                jCell.Value = rawJ
            End If
        Next jCell
    End If

    Dim kmRange As Range, kmCell As Range
    Set kmRange = Intersect(Target, Me.Range("K4:K23,M4:M23"))
    If Not kmRange Is Nothing Then
        For Each kmCell In kmRange.Cells
            Dim cleanedKM As String
            cleanedKM = NormalizeDecimal(CStr(kmCell.Value))
            If cleanedKM = "" Then
                kmCell.Value = 0
            Else
                kmCell.Value = Val(Replace(cleanedKM, ",", "."))
            End If
        Next kmCell
    End If

    If Not Intersect(Target, Me.Range("B16")) Is Nothing Then
        Me.Range("B16").Value = NormalizeRizInput(CStr(Me.Range("B16").Value))
    End If

    If Not Intersect(Target, Me.Range("B17")) Is Nothing Then
        Me.Range("B17").Value = NormalizeKInput(CStr(Me.Range("B17").Value))
    End If

    If Not Intersect(Target, Me.Range("L4")) Is Nothing Then
        Me.Range("L5:L23").Value = Me.Range("L4").Value
    End If

    If Not Intersect(Target, Me.Range("N4")) Is Nothing Then
        Me.Range("N5:N23").Value = Me.Range("N4").Value
    End If

SafeExit:
    mColorSignature = BuildColorSettingsSignature(Me)
    Me.Protect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)
    Application.EnableEvents = True
End Sub
'@

$historySheetCode = @'
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    If Target Is Nothing Then Exit Sub

    Dim lRange As Range
    Set lRange = Intersect(Target, Me.Columns(12))  ' Column L
    If lRange Is Nothing Then Exit Sub

    Application.EnableEvents = False
    Me.Unprotect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)

    Dim cell As Range
    For Each cell In lRange.Cells
        If cell.Locked Then GoTo NextLCell
        If cell.HasFormula Then GoTo NextLCell
        Dim cleaned As String
        cleaned = NormalizeDecimal(CStr(cell.Value))
        If cleaned = "" Then
            cell.Value = 0
        Else
            cell.Value = Val(Replace(cleaned, ",", "."))
        End If
NextLCell:
    Next cell

SafeExit:
    Me.Protect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)
    Application.EnableEvents = True
End Sub
'@

$mrsSheetCode = @'
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    If Target Is Nothing Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Row < 3 Then Exit Sub

    Dim col As Long
    col = Target.Column

    ' Column M (13) = duration, Column I (9) = pause rows only
    If col <> 13 And col <> 9 Then Exit Sub

    ' Skip formula cells (data row end times in col I)
    If Target.HasFormula Then Exit Sub

    Application.EnableEvents = False
    Me.Unprotect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)

    Dim cleaned As String
    cleaned = NormalizeDecimal(CStr(Target.Value))

    Dim numVal As Double
    If cleaned = "" Then
        numVal = 0
    Else
        numVal = Val(Replace(cleaned, ",", "."))
    End If

    Target.NumberFormat = "0.00"
    Target.Value = numVal
    Target.NumberFormat = "@"

SafeExit:
    Me.Protect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)
    Application.EnableEvents = True
End Sub
'@

$workbookCode = @'
Private Sub Workbook_Open()
    RefreshWorkbookColors
End Sub
'@

$excel = $null
$wb = $null
$securityKey = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$hadSecurityKey = Test-Path $securityKey
$previousAccessVBOM = $null
$hadAccessVBOMValue = $false
$restoreExpected = $null

try {
    if (-not $hadSecurityKey) {
        New-Item -Path $securityKey -Force | Out-Null
    }
    try {
        $previousAccessVBOM = (Get-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction Stop).AccessVBOM
        $hadAccessVBOMValue = $true
    } catch {
        $hadAccessVBOMValue = $false
    }
    New-ItemProperty -Path $securityKey -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $wb = $excel.Workbooks.Add()

    while ($wb.Worksheets.Count -lt 4) {
        $null = $wb.Worksheets.Add()
    }

    $wsInput = $wb.Worksheets.Item(1)
    $wsInput.Name = (RU @(1042,1074,1086,1076))
    $wsResult = $wb.Worksheets.Item(2)
    $wsResult.Name = (RU @(1056,1077,1079,1091,1083,1100,1090,1072,1090))
    $wsHistory = $wb.Worksheets.Item(3)
    $wsHistory.Name = (RU @(1048,1089,1090,1086,1088,1080,1103))
    $wsTechCards = $wb.Worksheets.Item(4)
    $wsTechCards.Name = (RU @(1055,1072,1088,1089,1080,1085,1075,32,77,82,83))

    # --- Title row ---
    $wsInput.Range("A1:O1").Merge() | Out-Null
    $wsInput.Range("A1").Value = (RU @(84,105,109,101,84,111,84,97,98,108,101,45,86,66,65,32,1043,1072,1083,1080,1084,1079,1103,1085,1086,1074,32,1043,46,1056,46,32,50,48,50,54))
    $wsInput.Range("A1").Font.Bold = $true
    $wsInput.Range("A1").Font.Size = 16
    $wsInput.Range("A1:O1").HorizontalAlignment = -4108  # xlCenter
    $wsInput.Range("A1:O1").VerticalAlignment = -4108     # xlCenter
    $wsInput.Rows(1).RowHeight = 50
    # --- Separator row: merged A2:O2 ---
    $wsInput.Range("A2:O2").Merge() | Out-Null

    $wsInput.Range("A3").Value = (RU @(1053,1086,1084,1077,1088,32,1047,1072,1082,1072,1079,1072))
    $wsInput.Range("A4").Value = (RU @(1053,1072,1080,1084,1077,1085,1086,1074,1072,1085,1080,1077))
    $wsInput.Range("A5").Value = (RU @(1044,1072,1090,1072,32,1085,1072,1095,1072,1083,1072))
    $wsInput.Range("A6").Value = (RU @(1042,1088,1077,1084,1103,32,1085,1072,1095,1072,1083,1072))
    $wsInput.Range("A7").Value = (RU @(1044,1072,1090,1072,32,1087,1088,1086,1074,1086,1076,1082,1080))
    $wsInput.Range("A8").Value = (RU @(1050,1086,1083,45,1074,1086,32,1086,1087,1077,1088,1072,1094,1080,1081))
    $wsInput.Range("A9").Value = (RU @(1050,1086,1083,45,1074,1086,32,1080,1089,1087,1086,1083,1085,1080,1090,1077,1083,1077,1081))
    $wsInput.Range("A10").Value = (RU @(1054,1073,1077,1076,32,49,32,40,1063,1063,58,1052,1052,41))
    $wsInput.Range("A11").Value = (RU @(1054,1073,1077,1076,32,50,32,40,48,48,58,48,48,32,45,32,1086,1090,1082,1083,46,41))
    $wsInput.Range("A12").Value = (RU @(1044,1083,1080,1090,1077,1083,1100,1085,1086,1089,1090,1100,32,1086,1073,1077,1076,1072,32,40,1084,1080,1085,41))
    $wsInput.Range("A13").Value = (RU @(90,55,58,32,1057,1086,1089,1090,1086,1103,1085,1080,1077,32,1076,1086,32,1088,1072,1073,1086,1090))
    $wsInput.Range("A14").Value = (RU @(90,55,58,32,1044,1086,1087,46,32,1088,1072,1073,1086,1090,1099))
    $wsInput.Range("A15").Value = (RU @(90,55,58,32,1056,1077,1082,1086,1084,1077,1085,1076,1072,1094,1080,1080))
    $wsInput.Range("A16").Value = (RU @(90,55,58,32,82,1080,1079,40,1052,1086,1084,41))
    $wsInput.Range("A17").Value2 = "Z7: K"

    $today = Get-Date
    $todayOA = [double]$today.ToOADate()
    $todayDateOnly = [Math]::Floor($todayOA)
    $wsInput.Range("B5").Value2 = $todayDateOnly
    $wsInput.Range("B5").NumberFormatLocal = [char]0x0414 + [char]0x0414 + "." + [char]0x041C + [char]0x041C + "." + [char]0x0413 + [char]0x0413 + [char]0x0413 + [char]0x0413
    $wsInput.Range("B6").Value2 = (8/24)
    $wsInput.Range("B6").NumberFormatLocal = [char]0x0447 + [char]0x0447 + ":" + [char]0x043C + [char]0x043C + ":" + [char]0x0441 + [char]0x0441
    $wsInput.Range("B7").Value2 = $todayDateOnly
    $wsInput.Range("B7").NumberFormatLocal = [char]0x0414 + [char]0x0414 + "." + [char]0x041C + [char]0x041C + "." + [char]0x0413 + [char]0x0413 + [char]0x0413 + [char]0x0413
    $wsInput.Range("B8").Value2 = 2
    $wsInput.Range("B9").Value2 = 2
    $wsInput.Range("B10").Value2 = "12:00"
    $wsInput.Range("B11").Value2 = "00:00"
    $wsInput.Range("B12").Value2 = 45
    $wsInput.Range("B13").Value = (RU @(1079,1072,1084,1077,1095,1072,1085,1080,1081,32,1085,1077,1090))
    $wsInput.Range("B14").Value = (RU @(1085,1077,1090))
    $wsInput.Range("B15").Value = (RU @(1085,1077,1090))
    $wsInput.Range("B16").Value2 = ""
    $wsInput.Range("B17").Value2 = ""
    $wsInput.Range("B16:B17").NumberFormat = "@"

    $wsInput.Range("D3").Value = (RU @(1048,1089,1087,1086,1083,1085,1080,1090,1077,1083,1100))
    $wsInput.Range("E3").Value = (RU @(8470))
    $wsInput.Range("D3:E3").Font.Bold = $true

    $activeWorkerCount = [int]$wsInput.Range("B9").Value2
    if ($activeWorkerCount -lt 1) { $activeWorkerCount = 1 }
    if ($activeWorkerCount -gt 10) { $activeWorkerCount = 10 }
    foreach ($i in 1..10) {
        $row = 3 + $i
        $wsInput.Cells.Item($row, 4).Value = ((RU @(1048,1089,1087,1086,1083,1085,1080,1090,1077,1083,1100)) + " " + $i)
        $wsInput.Cells.Item($row, 5).NumberFormat = "00000000"
        if ($i -le $activeWorkerCount) {
            $wsInput.Cells.Item($row, 4).Font.Color = 0
            $wsInput.Cells.Item($row, 5).Font.Color = 0
            $wsInput.Cells.Item($row, 5).Interior.Pattern = -4142
            $wsInput.Cells.Item($row, 5).Borders.LineStyle = 1
        } else {
            $wsInput.Cells.Item($row, 4).Font.Color = 16777215
            $wsInput.Cells.Item($row, 5).Font.Color = 16777215
            $wsInput.Cells.Item($row, 5).Interior.Color = 16777215
            $wsInput.Cells.Item($row, 5).Borders.LineStyle = -4142
        }
    }

    $wsInput.Range("G3").Value = (RU @(8470))
    $wsInput.Range("H3").Value = (RU @(1055,1044,1058,1042))
    $wsInput.Range("I3").Value = (RU @(1054,1087,1077,1088,1072,1094,1080,1103))
    $wsInput.Range("J3").Value = "-"
    $wsInput.Range("K3").Value = (RU @(1044,1083,1080,1090,1077,1083,1100,1085,1086,1089,1090,1100))
    $wsInput.Range("L3").Value = (RU @(1045,1076,1080,1085,1080,1094,1072,32,40,1084,1080,1085,47,1095,1072,1089,41))
    $wsInput.Range("M3").Value = (RU @(1055,1072,1091,1079,1072))
    $wsInput.Range("N3").Value = (RU @(1045,1076,1080,1085,1080,1094,1072,32,40,1084,1080,1085,47,1095,1072,1089,41))
    $wsInput.Range("O3").Value = (RU @(1059,1095,1072,1089,1090,1085,1080,1082,1080))
    $wsInput.Range("G3:O3").Font.Bold = $true

    $wsInput.Range("G4").Value2 = 1
    $wsInput.Range("I4").Value = (RU @(1054,1087,1077,1088,1072,1094,1080,1103,32,49))
    $wsInput.Range("K4").Value2 = 0
    $wsInput.Range("L4").Value = (RU @(1084,1080,1085))
    $wsInput.Range("N4").Value = (RU @(1084,1080,1085))

    $wsInput.Range("G5").Value2 = 2
    $wsInput.Range("I5").Value = (RU @(1054,1087,1077,1088,1072,1094,1080,1103,32,50))
    $wsInput.Range("K5").Value2 = 0
    $wsInput.Range("L5").Value = (RU @(1084,1080,1085))
    $wsInput.Range("N5").Value = (RU @(1084,1080,1085))

    # Default values: Пауза=0, Участники=пусто for active operations only
    for ($r = 4; $r -le 5; $r++) {
        $wsInput.Cells.Item($r, 13).Value2 = 0   # M = Пауза
        $wsInput.Cells.Item($r, 15).Value2 = ""  # O = Участники
    }

    # K4:K23, M4:M23: text format for safe input (NormalizeDecimal in VBA)
    $wsInput.Range("K4:K23").NumberFormat = "@"
    $wsInput.Range("M4:M23").NumberFormat = "@"

    $wsInput.Columns("A:A").ColumnWidth = 25
    $wsInput.Columns("B:B").ColumnWidth = 40
    $wsInput.Columns("C:C").ColumnWidth = 3
    $wsInput.Columns("D:D").ColumnWidth = 20
    $wsInput.Columns("E:E").ColumnWidth = 10
    $wsInput.Columns("F:F").ColumnWidth = 3
    $wsInput.Columns("G:G").ColumnWidth = 4
    $wsInput.Columns("H:H").ColumnWidth = 10
    $wsInput.Columns("I:I").ColumnWidth = 50
    $wsInput.Columns("J:J").ColumnWidth = 10
    $wsInput.Columns("K:K").ColumnWidth = 15
    $wsInput.Columns("L:L").ColumnWidth = 18
    $wsInput.Columns("M:M").ColumnWidth = 6
    $wsInput.Columns("N:N").ColumnWidth = 18
    $wsInput.Columns("O:O").ColumnWidth = 18
    $wsInput.Cells.HorizontalAlignment = -4108
    $wsInput.Cells.VerticalAlignment = -4108
    $wsInput.Cells.WrapText = $true

    # --- Light red for non-editable area ---
    $lightRed = 230*65536 + 230*256 + 255  # RGB(255,230,230) in BGR
    $wsInput.Cells.Interior.Color = $lightRed

    # --- Clear color on always-editable cells ---
    $alwaysEditable = @(
        "B3:B17",
        "L4","N4"
    )
    foreach ($addr in $alwaysEditable) {
        $wsInput.Range($addr).Interior.Pattern = -4142  # xlNone
        $wsInput.Range($addr).Locked = $false
        $wsInput.Range($addr).Borders.LineStyle = 1  # xlContinuous
    }

    # --- Workers: unlock only active rows (default 2) ---
    $defWorkers = 2
    for ($i = 1; $i -le 10; $i++) {
        $row = $i + 3
        $wsInput.Cells.Item($row, 4).Locked = $true
        if ($i -le $defWorkers) {
            $wsInput.Cells.Item($row, 5).Interior.Pattern = -4142
            $wsInput.Cells.Item($row, 5).Locked = $false
            $wsInput.Cells.Item($row, 5).Borders.LineStyle = 1
        } else {
            $wsInput.Cells.Item($row, 5).Interior.Color = $lightRed
            $wsInput.Cells.Item($row, 5).Locked = $true
            $wsInput.Cells.Item($row, 5).Borders.LineStyle = -4142
            $wsInput.Cells.Item($row, 5).Font.Color = $lightRed
            $wsInput.Cells.Item($row, 4).Font.Color = $lightRed
        }
    }

    # --- Operations: unlock only active rows (default 2) ---
    $defOps = 2
    # Editable op columns: H(8), I(9), J(10), K(11), M(13), O(15)
    $editOpCols = @(8, 9, 10, 11, 13, 15)
    # Synced locked columns: L(12), N(14)
    $syncOpCols = @(12, 14)
    for ($i = 1; $i -le 20; $i++) {
        $row = $i + 3
        if ($i -le $defOps) {
            foreach ($c in $editOpCols) {
                $wsInput.Cells.Item($row, $c).Interior.Pattern = -4142
                $wsInput.Cells.Item($row, $c).Locked = $false
                $wsInput.Cells.Item($row, $c).Borders.LineStyle = 1
            }
            foreach ($c in $syncOpCols) {
                $wsInput.Cells.Item($row, $c).Borders.LineStyle = 1
                if ($i -eq 1) {
                    # L4/N4 are editable source cells
                    $wsInput.Cells.Item($row, $c).Interior.Pattern = -4142
                    $wsInput.Cells.Item($row, $c).Locked = $false
                } else {
                    $wsInput.Cells.Item($row, $c).Interior.Color = $lightRed
                    $wsInput.Cells.Item($row, $c).Locked = $true
                }
            }
        } else {
            for ($c = 7; $c -le 15; $c++) {
                $wsInput.Cells.Item($row, $c).Interior.Color = $lightRed
                $wsInput.Cells.Item($row, $c).Font.Color = $lightRed
                $wsInput.Cells.Item($row, $c).Borders.LineStyle = -4142
                $wsInput.Cells.Item($row, $c).Locked = $true
            }
        }
    }

    # Lock M4 (pause for first op) — no history in fresh file
    $wsInput.Range("M4").Locked = $true
    $wsInput.Range("M4").Interior.Color = $lightRed

    # --- L5:L23, N5:N23: synced values ---
    $minText = RU @(1084,1080,1085)
    $wsInput.Range("L5:L23").Value = $minText
    $wsInput.Range("N5:N23").Value = $minText

    # --- Data validation with error messages ---
    $errTitle = RU @(1054,1096,1080,1073,1082,1072,32,1074,1074,1086,1076,1072)  # "Ошибка ввода"

    # B3: integer, max 12 digits (VBA normalizes)
    $wsInput.Range("B3").NumberFormat = "000000000000"
    $wsInput.Range("B3").Validation.Delete()
    $wsInput.Range("B3").Validation.Add(1, 1, 1, "0", "999999999999") | Out-Null
    $wsInput.Range("B3").Validation.IgnoreBlank = $true
    $wsInput.Range("B3").Validation.ErrorTitle = $errTitle
    $wsInput.Range("B3").Validation.ErrorMessage = (RU @(1052,1072,1082,1089,1080,1084,1091,1084,32,49,50,32,1094,1080,1092,1088))  # "Максимум 12 цифр"

    # B4: text, max 100 chars
    $wsInput.Range("B4").Validation.Delete()
    $wsInput.Range("B4").Validation.Add(6, 1, 8, "100") | Out-Null
    $wsInput.Range("B4").Validation.IgnoreBlank = $true
    $wsInput.Range("B4").Validation.ErrorTitle = $errTitle
    $wsInput.Range("B4").Validation.ErrorMessage = (RU @(1052,1072,1082,1089,1080,1084,1091,1084,32,49,48,48,32,1089,1080,1084,1074,1086,1083,1086,1074))  # "Максимум 100 символов"

    # B8: integer, 1-20 (operations)
    $wsInput.Range("B8").Validation.Delete()
    $wsInput.Range("B8").Validation.Add(1, 1, 1, "1", "20") | Out-Null
    $wsInput.Range("B8").Validation.ErrorTitle = $errTitle
    $wsInput.Range("B8").Validation.ErrorMessage = (RU @(1062,1077,1083,1086,1077,32,1095,1080,1089,1083,1086,32,1086,1090,32,49,32,1076,1086,32,50,48))  # "Целое число от 1 до 20"

    # B9: integer, 1-10 (workers)
    $wsInput.Range("B9").Validation.Delete()
    $wsInput.Range("B9").Validation.Add(1, 1, 1, "1", "10") | Out-Null
    $wsInput.Range("B9").Validation.ErrorTitle = $errTitle
    $wsInput.Range("B9").Validation.ErrorMessage = (RU @(1062,1077,1083,1086,1077,32,1095,1080,1089,1083,1086,32,1086,1090,32,49,32,1076,1086,32,49,48))  # "Целое число от 1 до 10"

    # B12: integer, 1-600
    $wsInput.Range("B12").Validation.Delete()
    $wsInput.Range("B12").Validation.Add(1, 1, 1, "1", "600") | Out-Null
    $wsInput.Range("B12").Validation.ErrorTitle = $errTitle
    $wsInput.Range("B12").Validation.ErrorMessage = (RU @(1062,1077,1083,1086,1077,32,1095,1080,1089,1083,1086,32,1086,1090,32,49,32,1076,1086,32,54,48,48))  # "Целое число от 1 до 600"

    # B13:B15: text, max 100 chars
    foreach ($cell in @("B13","B14","B15")) {
        $wsInput.Range($cell).Validation.Delete()
        $wsInput.Range($cell).Validation.Add(6, 1, 8, "100") | Out-Null
        $wsInput.Range($cell).Validation.IgnoreBlank = $true
        $wsInput.Range($cell).Validation.ErrorTitle = $errTitle
        $wsInput.Range($cell).Validation.ErrorMessage = (RU @(1052,1072,1082,1089,1080,1084,1091,1084,32,49,48,48,32,1089,1080,1084,1074,1086,1083,1086,1074))
    }

    # B16:B17: text, max 20 chars (VBA normalizes)
    foreach ($cell in @("B16","B17")) {
        $wsInput.Range($cell).Validation.Delete()
        $wsInput.Range($cell).Validation.Add(6, 1, 8, "20") | Out-Null
        $wsInput.Range($cell).Validation.IgnoreBlank = $true
        $wsInput.Range($cell).Validation.ErrorTitle = $errTitle
        $wsInput.Range($cell).Validation.ErrorMessage = (RU @(1052,1072,1082,1089,1080,1084,1091,1084,32,50,48,32,1089,1080,1084,1074,1086,1083,1086,1074))  # "Максимум 20 символов"
    }

    # E4:E13: integer, 8 digits
    $wsInput.Range("E4:E13").Validation.Delete()
    $wsInput.Range("E4:E13").Validation.Add(1, 1, 1, "0", "99999999") | Out-Null
    $wsInput.Range("E4:E13").Validation.IgnoreBlank = $true
    $wsInput.Range("E4:E13").Validation.ErrorTitle = $errTitle
    $wsInput.Range("E4:E13").Validation.ErrorMessage = (RU @(1052,1072,1082,1089,1080,1084,1091,1084,32,56,32,1094,1080,1092,1088))  # "Максимум 8 цифр"

    # H4:H23: integer, max 8 digits (VBA normalizes)
    $wsInput.Range("H4:H23").NumberFormat = "00000000"
    $wsInput.Range("H4:H23").Validation.Delete()
    $wsInput.Range("H4:H23").Validation.Add(1, 1, 1, "0", "99999999") | Out-Null
    $wsInput.Range("H4:H23").Validation.IgnoreBlank = $true
    $wsInput.Range("H4:H23").Validation.ErrorTitle = $errTitle
    $wsInput.Range("H4:H23").Validation.ErrorMessage = (RU @(1052,1072,1082,1089,1080,1084,1091,1084,32,56,32,1094,1080,1092,1088))  # "Максимум 8 цифр"

    # I4:I23: text, max 100 chars
    $wsInput.Range("I4:I23").Validation.Delete()
    $wsInput.Range("I4:I23").Validation.Add(6, 1, 8, "100") | Out-Null
    $wsInput.Range("I4:I23").Validation.IgnoreBlank = $true
    $wsInput.Range("I4:I23").Validation.ErrorTitle = $errTitle
    $wsInput.Range("I4:I23").Validation.ErrorMessage = (RU @(1052,1072,1082,1089,1080,1084,1091,1084,32,49,48,48,32,1089,1080,1084,1074,1086,1083,1086,1074))

    # J4:J23: decimal, 0-999999.99 (Data Validation)
    $decSep = [string]$excel.International(3)
    $maxDec = "999999" + $decSep + "99"
    $decErrMsg = (RU @(1063,1080,1089,1083,1086,32,1086,1090,32,48,32,1076,1086,32,57,57,57,57,57,57,44,57,57))  # "Число от 0 до 999999,99"
    foreach ($rng in @("J4:J23")) {
        $wsInput.Range($rng).Validation.Delete()
        $wsInput.Range($rng).Validation.Add(2, 1, 1, "0", $maxDec) | Out-Null
        $wsInput.Range($rng).Validation.IgnoreBlank = $true
        $wsInput.Range($rng).Validation.ErrorTitle = $errTitle
        $wsInput.Range($rng).Validation.ErrorMessage = $decErrMsg
    }
    # K4:K23, M4:M23: text format, normalized by NormalizeDecimal in VBA
    foreach ($rng in @("K4:K23","M4:M23")) {
        $wsInput.Range($rng).Validation.Delete()
    }

    # L4, N4: selector мин/час (only first row editable)
    $listSep = $excel.International(5)
    $unitList = (RU @(1084,1080,1085)) + $listSep + (RU @(1095,1072,1089))
    $unitErrMsg = (RU @(1042,1099,1073,1077,1088,1080,1090,1077,32,1084,1080,1085,32,1080,1083,1080,32,1095,1072,1089))  # "Выберите мин или час"
    foreach ($rng in @("L4","N4")) {
        $wsInput.Range($rng).Validation.Delete()
        $wsInput.Range($rng).Validation.Add(3, 1, 1, $unitList) | Out-Null
        $wsInput.Range($rng).Validation.IgnoreBlank = $true
        $wsInput.Range($rng).Validation.InCellDropdown = $true
        $wsInput.Range($rng).Validation.ErrorTitle = $errTitle
        $wsInput.Range($rng).Validation.ErrorMessage = $unitErrMsg
    }

    # O4:O23: text, max 100 chars
    $wsInput.Range("O4:O23").Validation.Delete()
    $wsInput.Range("O4:O23").Validation.Add(6, 1, 8, "100") | Out-Null
    $wsInput.Range("O4:O23").Validation.IgnoreBlank = $true
    $wsInput.Range("O4:O23").Validation.ErrorTitle = $errTitle
    $wsInput.Range("O4:O23").Validation.ErrorMessage = (RU @(1052,1072,1082,1089,1080,1084,1091,1084,32,49,48,48,32,1089,1080,1084,1074,1086,1083,1086,1074))

    $wsResult.Range("A1").Value2 = ""
    $wsResult.Range("B1").Value = (RU @(8470))
    $wsResult.Range("C1").Value = (RU @(1054,1087,1077,1088,1072,1094,1080,1103))
    $wsResult.Range("D1").Value = (RU @(1054,1073,1077,1076,63))
    $wsResult.Range("E1").Value = (RU @(1055,1072,1091,1079,1072))
    $wsResult.Range("F1").Value = (RU @(1056,1072,1073,1086,1090,1072,32,40,1095,1072,1089,41))
    $wsResult.Range("G1").Value = (RU @(1055,1044,1058,1042))
    $wsResult.Range("H1").Value2 = "-"
    $wsResult.Range("I1").Value2 = "-"
    $wsResult.Range("J1").Value2 = "-"
    $wsResult.Range("K1").Value2 = "-"
    $wsResult.Range("L1").Value = (RU @(1056,1072,1073,1086,1090,1072,32,40,1084,1080,1085,41))
    $wsResult.Range("M1").Value2 = "-"
    $wsResult.Range("N1").Value2 = "-"
    $wsResult.Range("O1").Value = (RU @(1044,1072,1090,1072,32,1087,1088,1086,1074,1086,1076,1082,1080))
    $wsResult.Range("P1").Value = (RU @(1048,1089,1087,1086,1083,1085,1080,1090,1077,1083,1100))
    $wsResult.Range("Q1").Value2 = "-"
    $wsResult.Range("R1").Value = (RU @(1044,1072,1090,1072,32,1053,1072,1095,1072,1083,1072))
    $wsResult.Range("S1").Value = (RU @(1042,1088,1077,1084,1103,32,1053,1072,1095,1072,1083,1072))
    $wsResult.Range("T1").Value = (RU @(1044,1072,1090,1072,32,1050,1086,1085,1094,1072))
    $wsResult.Range("U1").Value = (RU @(1042,1088,1077,1084,1103,32,1050,1086,1085,1094,1072))
    $wsResult.Range("V1").Value2 = "INDEX"
    $wsResult.Range("A1:V1").Font.Bold = $true
    $wsResult.Columns("A:V").AutoFit() | Out-Null
    $wsResult.Columns("A:A").ColumnWidth = 2.7
    $wsResult.Columns("H:K").ColumnWidth = 2.7
    $wsResult.Columns("M:N").ColumnWidth = 2.7
    $wsResult.Columns("Q:Q").ColumnWidth = 2.7

    # --- Result sheet formatting ---
    $lightRed = 230*65536 + 230*256 + 255  # RGB(255,230,230) in BGR
    $wsResult.Cells.Interior.Color = $lightRed
    $wsResult.Cells.HorizontalAlignment = -4108  # xlCenter
    $wsResult.Cells.VerticalAlignment = -4108    # xlCenter
    $wsResult.Cells.Locked = $true
    # Borders on header row
    $wsResult.Range("B1:V1").Borders.LineStyle = 1  # xlContinuous
    $wsResult.Protect((RU 49,49,52,55,48,57), $true, $true, $false, $false)

    $wsHistory.Range("B1:V1").Merge() | Out-Null
    $wsHistory.Range("B1").Value = (RU @(1048,1089,1090,1086,1088,1080,1103))
    $wsHistory.Range("B1").Font.Bold = $true
    $wsHistory.Range("B1").Font.Size = 14
    $wsHistory.Range("B1:V1").Borders.LineStyle = 1  # xlContinuous
    $wsHistory.Cells.Font.Size = 14
    $wsHistory.Columns("A:A").ColumnWidth = 2.7
    $wsHistory.Columns("B:B").ColumnWidth = 3.5
    $wsHistory.Columns("C:C").ColumnWidth = 60
    $wsHistory.Columns("D:D").ColumnWidth = 8
    $wsHistory.Columns("E:E").ColumnWidth = 11
    $wsHistory.Columns("F:F").ColumnWidth = 15
    $wsHistory.Columns("G:G").ColumnWidth = 12
    $wsHistory.Columns("H:K").ColumnWidth = 1.5
    $wsHistory.Columns("L:L").ColumnWidth = 16
    $wsHistory.Columns("M:N").ColumnWidth = 1.5
    $wsHistory.Columns("O:O").ColumnWidth = 18
    $wsHistory.Columns("P:P").ColumnWidth = 16
    $wsHistory.Columns("Q:Q").ColumnWidth = 1.5
    $wsHistory.Columns("R:R").ColumnWidth = 16
    $wsHistory.Columns("S:S").ColumnWidth = 16
    $wsHistory.Columns("T:T").ColumnWidth = 16
    $wsHistory.Columns("U:U").ColumnWidth = 16
    $wsHistory.Columns("V:V").ColumnWidth = 10

    # --- History sheet formatting (same as Result) ---
    $wsHistory.Cells.Interior.Color = $lightRed
    $wsHistory.Cells.HorizontalAlignment = -4108  # xlCenter
    $wsHistory.Cells.VerticalAlignment = -4108    # xlCenter
    $wsHistory.Cells.Locked = $true
    $wsHistory.Protect((RU 49,49,52,55,48,57), $true, $true, $false, $false)

    # --- Setup MRS sheet ---
    $wsTechCards.Cells.Interior.Color = $lightRed
    $wsTechCards.Cells.HorizontalAlignment = -4108  # xlCenter
    $wsTechCards.Cells.VerticalAlignment = -4108     # xlCenter
    $wsTechCards.Cells.Font.Size = 14
    $wsTechCards.Cells.Locked = $true

    # Title row 1: "Парсинг MRS" merged B1:N1
    $wsTechCards.Range("B1:N1").Merge()
    $wsTechCards.Range("B1").Value = (RU @(1055,1072,1088,1089,1080,1085,1075,32,77,82,83))  # Парсинг MRS
    $wsTechCards.Range("B1").Font.Bold = $true
    $wsTechCards.Range("B1:N1").Borders.LineStyle = 1

    # Row 2 = empty separator
    # Column headers are written per-block in VBA LoadMRS

    # Column widths
    $wsTechCards.Columns("A:A").ColumnWidth = 2.7
    $wsTechCards.Columns("B:B").ColumnWidth = 4
    $wsTechCards.Columns("C:C").ColumnWidth = 13
    $wsTechCards.Columns("D:D").ColumnWidth = 45
    $wsTechCards.Columns("E:E").ColumnWidth = 12
    $wsTechCards.Columns("F:F").ColumnWidth = 15
    $wsTechCards.Columns("G:G").ColumnWidth = 17
    $wsTechCards.Columns("H:H").ColumnWidth = 15
    $wsTechCards.Columns("I:I").ColumnWidth = 17
    $wsTechCards.Columns("J:J").ColumnWidth = 24
    $wsTechCards.Columns("K:K").ColumnWidth = 24
    $wsTechCards.Columns("L:L").ColumnWidth = 24
    $wsTechCards.Columns("M:M").ColumnWidth = 24
    $wsTechCards.Columns("N:N").ColumnWidth = 10
    $wsTechCards.Protect((RU 49,49,52,55,48,57), $true, $true, $false, $false)

    $buttonLeft = [double]$wsInput.Range("A19").Left
    $buttonTop = [double]$wsInput.Range("A19").Top
    $buttonWidth = [double]$wsInput.Columns("A:A").Width + [double]$wsInput.Columns("B:B").Width
    $button = $wsInput.Shapes.AddShape(1, $buttonLeft, $buttonTop, $buttonWidth, 32)
    $button.Name = "btnGenerateSchedule"
    $button.TextFrame.Characters().Text = (RU @(1057,1092,1086,1088,1084,1080,1088,1086,1074,1072,1090,1100,32,1058,1072,1073,1083,1080,1094,1091))
    $button.TextFrame.HorizontalAlignment = -4108  # xlCenter
    $button.TextFrame.VerticalAlignment = -4108     # xlCenter
    $button.Fill.ForeColor.RGB = 5296274
    $button.Line.ForeColor.RGB = 3355443
    $button.OnAction = "GenerateAndAppendHistory"

    $clearButtonLeft = $buttonLeft
    $clearButtonTop = $buttonTop + 38
    $clearButton = $wsInput.Shapes.AddShape(1, $clearButtonLeft, $clearButtonTop, $buttonWidth, 28)
    $clearButton.Name = "btnClearSchedule"
    $clearButton.TextFrame.Characters().Text = (RU @(1054,1095,1080,1089,1090,1080,1090,1100,32,1048,1089,1090,1086,1088,1080,1102))
    $clearButton.TextFrame.HorizontalAlignment = -4108  # xlCenter
    $clearButton.TextFrame.VerticalAlignment = -4108     # xlCenter
    $clearButton.Fill.ForeColor.RGB = 10066329
    $clearButton.Line.ForeColor.RGB = 3355443
    $clearButton.OnAction = "ClearResultAndHistory"

    $mrsButtonLeft = $buttonLeft
    $mrsButtonTop = $clearButtonTop + 40
    $mrsButton = $wsInput.Shapes.AddShape(1, $mrsButtonLeft, $mrsButtonTop, $buttonWidth, 28)
    $mrsButton.Name = "btnLoadMRS"
    $mrsButton.TextFrame.Characters().Text = (RU @(1047,1072,1075,1088,1091,1079,1080,1090,1100,32,1044,1072,1085,1085,1099,1077,32,1048,1079,32,83,65,80))
    $mrsButton.TextFrame.HorizontalAlignment = -4108  # xlCenter
    $mrsButton.TextFrame.VerticalAlignment = -4108     # xlCenter
    $mrsButton.Fill.ForeColor.RGB = 12419407  # RGB(79,129,189) blue
    $mrsButton.Line.ForeColor.RGB = 3355443
    $mrsButton.TextFrame.Characters().Font.Color = 16777215  # white
    $mrsButton.OnAction = "LoadMRS"

    $clearMrsLeft = $buttonLeft
    $clearMrsTop = $mrsButtonTop + 34
    $clearMrsBtn = $wsInput.Shapes.AddShape(1, $clearMrsLeft, $clearMrsTop, $buttonWidth, 28)
    $clearMrsBtn.Name = "btnClearMRS"
    $clearMrsBtn.TextFrame.Characters().Text = (RU @(1054,1095,1080,1089,1090,1080,1090,1100,32,1055,1072,1088,1089,1080,1085,1075,32,77,82,83))
    $clearMrsBtn.TextFrame.HorizontalAlignment = -4108  # xlCenter
    $clearMrsBtn.TextFrame.VerticalAlignment = -4108     # xlCenter
    $clearMrsBtn.Fill.ForeColor.RGB = 10066329  # gray
    $clearMrsBtn.Line.ForeColor.RGB = 3355443
    $clearMrsBtn.OnAction = "ClearMRS"

    # --- Color settings (rows 30-36) ---
    $wsInput.Range("A30:B30").Merge() | Out-Null
    $wsInput.Range("A30").Value = (RU @(1053,1072,1089,1090,1088,1086,1081,1082,1080,32,1062,1074,1077,1090,1086,1074))  # "Настройки Цветов"
    $wsInput.Range("A30").Font.Bold = $true

    $wsInput.Cells.Item(31, 1).Value = (RU @(1047,1072,1073,1083,1086,1082,1080,1088,1086,1074,1072,1085,1085,1099,1077))                         # "Заблокированные"
    $wsInput.Cells.Item(32, 1).Value = (RU @(1056,1077,1076,1072,1082,1090,1080,1088,1091,1077,1084,1099,1077))                                   # "Редактируемые"
    $wsInput.Cells.Item(33, 1).Value = (RU @(1055,1072,1088,1089,1080,1085,1075,32,77,82,83,32,1044,1072,1090,1072))                              # "Парсинг MRS Дата"
    $wsInput.Cells.Item(34, 1).Value = (RU @(1055,1072,1088,1089,1080,1085,1075,32,77,82,83,32,1041,1088,1080,1075,1072,1076,1072))               # "Парсинг MRS Бригада"
    $wsInput.Cells.Item(35, 1).Value = (RU @(1055,1072,1088,1089,1080,1085,1075,32,77,82,83,32,1047,1072,1082,1072,1079))                         # "Парсинг MRS Заказ"
    $wsInput.Cells.Item(36, 1).Value = (RU @(1062,1074,1077,1090,32,1058,1077,1082,1089,1090,1072))                                               # "Цвет Текста"

    # Lock merged header row 30
    $wsInput.Range("A30:B30").Locked = $true
    # Lock label cells (A31:A36), unlock color sample cells (B31:B36)
    for ($r = 31; $r -le 36; $r++) {
        $wsInput.Cells.Item($r, 1).Locked = $true
        $wsInput.Cells.Item($r, 2).Locked = $false
    }

    # Set default fill colors and borders in B column
    $clrDefLocked    = 230*65536 + 230*256 + 255   # RGB(255,230,230) - same as $lightRed
    $clrDefMrsHeader = 231*65536 + 198*256 + 180   # RGB(180,198,231)
    $clrDefMrsSub    = 255*65536 + 220*256 + 200   # RGB(200,220,255)
    $clrDefMrsOrder  = 200*65536 + 235*256 + 200   # RGB(200,235,200)
    $clrDefFont      = 0                           # RGB(0,0,0) black

    $wsInput.Cells.Item(31, 2).Interior.Color = $clrDefLocked
    # Row 32 (Editable) - no fill by default (xlNone)
    $wsInput.Cells.Item(32, 2).Interior.Pattern = -4142  # xlNone
    $wsInput.Cells.Item(33, 2).Interior.Color = $clrDefMrsHeader
    $wsInput.Cells.Item(34, 2).Interior.Color = $clrDefMrsSub
    $wsInput.Cells.Item(35, 2).Interior.Color = $clrDefMrsOrder
    $wsInput.Cells.Item(36, 2).Interior.Color = $clrDefFont

    # Add borders to B31:B36
    for ($r = 31; $r -le 36; $r++) {
        $wsInput.Cells.Item($r, 2).Borders.LineStyle = 1  # xlContinuous
    }

    # --- Small "pick color" buttons in column C for rows 31-36 ---
    $pickBtnNames = @(
        @{ Row = 31; Action = "PickColorLocked" },
        @{ Row = 32; Action = "PickColorEditable" },
        @{ Row = 33; Action = "PickColorMrsHeader" },
        @{ Row = 34; Action = "PickColorMrsSub" },
        @{ Row = 35; Action = "PickColorMrsOrder" },
        @{ Row = 36; Action = "PickColorFont" }
    )
    foreach ($btn in $pickBtnNames) {
        $r = $btn.Row
        $btnLeft   = [double]$wsInput.Cells.Item($r, 3).Left
        $btnTop    = [double]$wsInput.Cells.Item($r, 3).Top + 1
        $btnHeight = [double]$wsInput.Rows($r).RowHeight - 2
        $btnWidth  = 20
        $shape = $wsInput.Shapes.AddShape(1, $btnLeft, $btnTop, $btnWidth, $btnHeight)
        $shape.Name = "btn_" + $btn.Action
        $shape.TextFrame.Characters().Text = "..."
        $shape.TextFrame.Characters().Font.Size = 8
        $shape.TextFrame.HorizontalAlignment = -4108  # xlCenter
        $shape.TextFrame.VerticalAlignment = -4108     # xlCenter
        $shape.Fill.ForeColor.RGB = 15790320  # light gray
        $shape.Line.ForeColor.RGB = 10066329  # darker gray
        $shape.OnAction = $btn.Action
    }

    # --- Protect input sheet (empty password) ---
    $wsInput.Protect((RU 49,49,52,55,48,57), $true, $true, $false, $false)

    try {
        $vbComp = $wb.VBProject.VBComponents.Add(1)
        $vbComp.Name = "modTimeToTable"
        $null = $vbComp.CodeModule.AddFromString($vbaCode)
        $sheetComp = $wb.VBProject.VBComponents.Item($wsInput.CodeName)
        $null = $sheetComp.CodeModule.AddFromString($inputSheetCode)
        $histComp = $wb.VBProject.VBComponents.Item($wsHistory.CodeName)
        $null = $histComp.CodeModule.AddFromString($historySheetCode)
        $mrsComp = $wb.VBProject.VBComponents.Item($wsTechCards.CodeName)
        $null = $mrsComp.CodeModule.AddFromString($mrsSheetCode)
        $wbComp = $wb.VBProject.VBComponents.Item($wb.CodeName)
        $null = $wbComp.CodeModule.AddFromString($workbookCode)
    } catch {
        throw "Failed to inject VBA code. In Excel enable: Trust Center -> Trust access to the VBA project object model. Details: $($_.Exception.Message)"
    }

    $xlOpenXMLWorkbookMacroEnabled = 52
    $targetDir = Split-Path -Path $OutputPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($targetDir) -and -not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir | Out-Null
    }
    $excel.EnableEvents = $true
    $wb.SaveAs($OutputPath, $xlOpenXMLWorkbookMacroEnabled)

    Write-Output "CREATED: $OutputPath"
} finally {
    try {
        if ($hadAccessVBOMValue) {
            $restoreExpected = [int]$previousAccessVBOM
            New-ItemProperty -Path $securityKey -Name AccessVBOM -Value ([int]$previousAccessVBOM) -PropertyType DWord -Force | Out-Null
        } else {
            $restoreExpected = $null
            Remove-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue
            if (-not $hadSecurityKey) {
                Remove-Item -Path $securityKey -Force -ErrorAction SilentlyContinue
            }
        }
    } catch {}

    if ($wb -ne $null) {
        try { $wb.Close($false) } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    $actualAccessVBOM = $null
    $hasAccessVBOMNow = $false
    try {
        $actualAccessVBOM = (Get-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction Stop).AccessVBOM
        $hasAccessVBOMNow = $true
    } catch {
        $hasAccessVBOMNow = $false
    }

    if ($restoreExpected -eq $null) {
        if ($hasAccessVBOMNow) {
            throw "Security rollback failed: AccessVBOM should be absent, but current value is '$actualAccessVBOM'."
        }
    } else {
        if (-not $hasAccessVBOMNow -or [int]$actualAccessVBOM -ne [int]$restoreExpected) {
            throw "Security rollback failed: AccessVBOM should be '$restoreExpected', but current value is '$actualAccessVBOM'."
        }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
