param(
    [string]$OutputPath = (Join-Path (Split-Path $PSScriptRoot -Parent) "TimeToTable_VBA.xlsm")
)

$ErrorActionPreference = "Stop"

function RU([int[]]$codes) {
    return -join ($codes | ForEach-Object { [char]$_ })
}

$vbaCode = @'
Option Explicit

Private Const CLR_ROW_LOCKED As Long = 31
Private Const CLR_ROW_EDITABLE As Long = 32
Private Const CLR_ROW_MRS_HEADER As Long = 33
Private Const CLR_ROW_MRS_SUBHEADER As Long = 34
Private Const CLR_ROW_MRS_ORDER As Long = 35
Private Const CLR_ROW_MRS_ORDER_UNCONF As Long = 36
Private Const CLR_ROW_HEADER As Long = 37
Private Const CLR_COL As Long = 2

Private Const CLR_DEF_LOCKED As Long = 15132415
Private Const CLR_DEF_MRS_HEADER As Long = 15189684
Private Const CLR_DEF_MRS_SUB As Long = 16768200
Private Const CLR_DEF_MRS_ORDER As Long = 13167560
Private Const CLR_DEF_MRS_ORDER_UNCONF As Long = 14277081
Private Const CLR_DEF_HEADER As Long = 13167560

Private Function ReadCellColor(ByVal cell As Range, ByVal defaultColor As Long) As Long
    If cell.Interior.Pattern = xlNone Or cell.Interior.Color = 0 Then
        ReadCellColor = defaultColor
    Else
        ReadCellColor = cell.Interior.Color
    End If
End Function

Public Function GetContrastColor(ByVal bgColor As Long) As Long
    If bgColor = xlNone Or bgColor = -4142 Then
        GetContrastColor = 0
        Exit Function
    End If
    
    Dim R As Long, G As Long, B As Long
    Dim luminance As Double
    
    R = bgColor Mod 256
    G = (bgColor \ 256) Mod 256
    B = (bgColor \ 65536) Mod 256
    
    luminance = (R * 299& + G * 587& + B * 114&) / 1000#
    
    If luminance > 128 Then
        GetContrastColor = 0
    Else
        GetContrastColor = 16777215
    End If
End Function

Private Sub ReadAllColors(ByVal wsIn As Worksheet, _
    ByRef clrLocked As Long, _
    ByRef clrEditHasColor As Boolean, ByRef clrEditable As Long, _
    ByRef clrMrsHeader As Long, ByRef clrMrsSub As Long, _
    ByRef clrMrsOrder As Long, ByRef clrMrsOrderUnconf As Long, _
    ByRef clrHeader As Long)

    clrLocked = ReadCellColor(wsIn.Cells(CLR_ROW_LOCKED, CLR_COL), CLR_DEF_LOCKED)

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
    clrMrsOrderUnconf = ReadCellColor(wsIn.Cells(CLR_ROW_MRS_ORDER_UNCONF, CLR_COL), CLR_DEF_MRS_ORDER_UNCONF)
    clrHeader = ReadCellColor(wsIn.Cells(CLR_ROW_HEADER, CLR_COL), CLR_DEF_HEADER)
End Sub

Public Sub GetOrderColors(ByRef clrMrsOrder As Long, ByRef clrMrsOrderUnconf As Long, Optional ByRef clrLocked As Long, Optional ByRef clrEditHasColor As Boolean, Optional ByRef clrEditable As Long)
    Dim wsIn As Worksheet
    Set wsIn = ThisWorkbook.Worksheets(2)
    clrMrsOrder = ReadCellColor(wsIn.Cells(CLR_ROW_MRS_ORDER, CLR_COL), CLR_DEF_MRS_ORDER)
    clrMrsOrderUnconf = ReadCellColor(wsIn.Cells(CLR_ROW_MRS_ORDER_UNCONF, CLR_COL), CLR_DEF_MRS_ORDER_UNCONF)
    clrLocked = ReadCellColor(wsIn.Cells(CLR_ROW_LOCKED, CLR_COL), CLR_DEF_LOCKED)
    If wsIn.Cells(CLR_ROW_EDITABLE, CLR_COL).Interior.Pattern = xlNone Then
        clrEditHasColor = False
        clrEditable = 0
    Else
        clrEditHasColor = True
        clrEditable = wsIn.Cells(CLR_ROW_EDITABLE, CLR_COL).Interior.Color
    End If
End Sub

Public Sub ColorOrderRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lastCol As Long, _
    ByVal clrOrder As Long, ByVal clrUnconf As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(r, 2), ws.Cells(r, lastCol))
    Dim v As String
    v = Trim$(CStr(ws.Cells(r, 2).Value))
    If v = UW(1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1086) Then
        rng.Interior.Color = clrOrder
        rng.Font.Color = GetContrastColor(clrOrder)
    ElseIf v = UW(1053, 1045, 32, 1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1086) Then
        rng.Interior.Color = clrUnconf
        rng.Font.Color = GetContrastColor(clrUnconf)
    End If
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

    wsIn.Cells(29, 1).Value = UW(1053, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1080, 32, 1062, 1074, 1077, 1090, 1086, 1074)
    On Error Resume Next
    wsIn.Range("A29:B29").Merge
    wsIn.Cells(29, 1).Font.Bold = True
    On Error GoTo 0

    wsIn.Cells(CLR_ROW_LOCKED, 1).Value = UW(1047, 1072, 1073, 1083, 1086, 1082, 1080, 1088, 1086, 1074, 1072, 1085, 1085, 1099, 1077)
    wsIn.Cells(CLR_ROW_EDITABLE, 1).Value = UW(1056, 1077, 1076, 1072, 1082, 1090, 1080, 1088, 1091, 1077, 1084, 1099, 1077)
    wsIn.Cells(CLR_ROW_MRS_HEADER, 1).Value = UW(1055, 1072, 1088, 1089, 1080, 1085, 1075, 32, 77, 82, 83, 32, 1044, 1072, 1090, 1072)
    wsIn.Cells(CLR_ROW_MRS_SUBHEADER, 1).Value = UW(1055, 1072, 1088, 1089, 1080, 1085, 1075, 32, 77, 82, 83, 32, 1041, 1088, 1080, 1075, 1072, 1076, 1072)
    wsIn.Cells(CLR_ROW_MRS_ORDER, 1).Value = UW(1047, 1072, 1082, 1072, 1079, 32, 1055, 1044, 1058, 1042)
    wsIn.Cells(CLR_ROW_MRS_ORDER_UNCONF, 1).Value = UW(1047, 1072, 1082, 1072, 1079, 32, 1053, 1045, 32, 1055, 1044, 1058, 1042)
    wsIn.Cells(CLR_ROW_HEADER, 1).Value = UW(1064, 1072, 1087, 1082, 1072)

    For r = CLR_ROW_LOCKED To CLR_ROW_HEADER
        wsIn.Cells(r, 1).Locked = True
        wsIn.Cells(r, CLR_COL).Locked = False
    Next r

    If wsIn.Cells(CLR_ROW_LOCKED, CLR_COL).Interior.Pattern = xlNone Or wsIn.Cells(CLR_ROW_LOCKED, CLR_COL).Interior.Color = 0 Then
        wsIn.Cells(CLR_ROW_LOCKED, CLR_COL).Interior.Color = CLR_DEF_LOCKED
    End If
    If wsIn.Cells(CLR_ROW_EDITABLE, CLR_COL).Interior.Pattern <> xlNone And wsIn.Cells(CLR_ROW_EDITABLE, CLR_COL).Interior.Color = 0 Then
        wsIn.Cells(CLR_ROW_EDITABLE, CLR_COL).Interior.Pattern = xlNone
    End If
    If wsIn.Cells(CLR_ROW_MRS_HEADER, CLR_COL).Interior.Pattern = xlNone Or wsIn.Cells(CLR_ROW_MRS_HEADER, CLR_COL).Interior.Color = 0 Then
        wsIn.Cells(CLR_ROW_MRS_HEADER, CLR_COL).Interior.Color = CLR_DEF_MRS_HEADER
    End If
    If wsIn.Cells(CLR_ROW_MRS_SUBHEADER, CLR_COL).Interior.Pattern = xlNone Or wsIn.Cells(CLR_ROW_MRS_SUBHEADER, CLR_COL).Interior.Color = 0 Then
        wsIn.Cells(CLR_ROW_MRS_SUBHEADER, CLR_COL).Interior.Color = CLR_DEF_MRS_SUB
    End If
    If wsIn.Cells(CLR_ROW_MRS_ORDER, CLR_COL).Interior.Pattern = xlNone Or wsIn.Cells(CLR_ROW_MRS_ORDER, CLR_COL).Interior.Color = 0 Then
        wsIn.Cells(CLR_ROW_MRS_ORDER, CLR_COL).Interior.Color = CLR_DEF_MRS_ORDER
    End If
    If wsIn.Cells(CLR_ROW_MRS_ORDER_UNCONF, CLR_COL).Interior.Pattern = xlNone Or wsIn.Cells(CLR_ROW_MRS_ORDER_UNCONF, CLR_COL).Interior.Color = 0 Then
        wsIn.Cells(CLR_ROW_MRS_ORDER_UNCONF, CLR_COL).Interior.Color = CLR_DEF_MRS_ORDER_UNCONF
    End If
    If wsIn.Cells(CLR_ROW_HEADER, CLR_COL).Interior.Pattern = xlNone Or wsIn.Cells(CLR_ROW_HEADER, CLR_COL).Interior.Color = 0 Then
        wsIn.Cells(CLR_ROW_HEADER, CLR_COL).Interior.Color = CLR_DEF_HEADER
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

Private Sub ApplyLockedStyle(ByVal rng As Range, ByVal clrLocked As Long)
    rng.Interior.Color = clrLocked
    rng.Font.Color = GetContrastColor(clrLocked)
End Sub

Private Sub ApplyEditableStyle(ByVal rng As Range, ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long)
    If clrEditHasColor Then
        rng.Interior.Color = clrEditable
        rng.Font.Color = GetContrastColor(clrEditable)
    Else
        rng.Interior.Pattern = xlNone
        rng.Font.Color = 0
    End If
End Sub

Public Function BuildColorSettingsSignature(ByVal wsIn As Worksheet) As String
    Dim r As Long
    Dim cell As Range

    For r = CLR_ROW_LOCKED To CLR_ROW_HEADER
        Set cell = wsIn.Cells(r, CLR_COL)
        BuildColorSettingsSignature = BuildColorSettingsSignature & "|" & _
            CStr(cell.Interior.Pattern) & ":" & _
            CStr(cell.Interior.Color)
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

Private Sub SyncPauseInputCell(ByVal wsIn As Worksheet, ByVal wsHist As Worksheet, _
    ByVal clrLocked As Long, _
    ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long)

    Dim histLastRow As Long
    histLastRow = wsHist.Cells(wsHist.Rows.Count, 2).End(xlUp).Row

    If histLastRow <= 3 Then
        wsIn.Cells(4, 14).Locked = True
        ApplyLockedStyle wsIn.Cells(4, 14), clrLocked
    Else
        wsIn.Cells(4, 14).Locked = False
        ApplyEditableStyle wsIn.Cells(4, 14), clrEditHasColor, clrEditable
    End If
    wsIn.Cells(4, 14).Borders.LineStyle = xlContinuous
End Sub

Private Sub RefreshInputSheetColors(ByVal wsIn As Worksheet, ByVal wsHist As Worksheet, _
    ByVal clrLocked As Long, _
    ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long, _
    ByVal clrMrsHeader As Long, ByVal clrMrsSub As Long, _
    ByVal clrMrsOrder As Long, ByVal clrMrsOrderUnconf As Long, _
    ByVal clrHeader As Long)

    Dim workerCount As Long, opCount As Long

    wsIn.Cells.Interior.Color = clrLocked
    wsIn.Cells.Font.Color = GetContrastColor(clrLocked)

    ApplyEditableStyle wsIn.Range("B3:B17"), clrEditHasColor, clrEditable
    wsIn.Range("B3:B17").Borders.LineStyle = xlContinuous
    ApplyEditableStyle wsIn.Range("L4"), clrEditHasColor, clrEditable
    ApplyEditableStyle wsIn.Range("M4"), clrEditHasColor, clrEditable
    ApplyEditableStyle wsIn.Range("O4"), clrEditHasColor, clrEditable
    wsIn.Range("L4").Borders.LineStyle = xlContinuous
    wsIn.Range("M4").Borders.LineStyle = xlContinuous
    wsIn.Range("O4").Borders.LineStyle = xlContinuous

    wsIn.Cells(CLR_ROW_LOCKED, CLR_COL).Interior.Color = clrLocked
    wsIn.Cells(CLR_ROW_LOCKED, CLR_COL).Font.Color = GetContrastColor(clrLocked)
    If clrEditHasColor Then
        wsIn.Cells(CLR_ROW_EDITABLE, CLR_COL).Interior.Color = clrEditable
        wsIn.Cells(CLR_ROW_EDITABLE, CLR_COL).Font.Color = GetContrastColor(clrEditable)
    Else
        wsIn.Cells(CLR_ROW_EDITABLE, CLR_COL).Interior.Pattern = xlNone
        wsIn.Cells(CLR_ROW_EDITABLE, CLR_COL).Font.Color = 0
    End If
    wsIn.Cells(CLR_ROW_MRS_HEADER, CLR_COL).Interior.Color = clrMrsHeader
    wsIn.Cells(CLR_ROW_MRS_HEADER, CLR_COL).Font.Color = GetContrastColor(clrMrsHeader)
    wsIn.Cells(CLR_ROW_MRS_SUBHEADER, CLR_COL).Interior.Color = clrMrsSub
    wsIn.Cells(CLR_ROW_MRS_SUBHEADER, CLR_COL).Font.Color = GetContrastColor(clrMrsSub)
    wsIn.Cells(CLR_ROW_MRS_ORDER, CLR_COL).Interior.Color = clrMrsOrder
    wsIn.Cells(CLR_ROW_MRS_ORDER, CLR_COL).Font.Color = GetContrastColor(clrMrsOrder)
    wsIn.Cells(CLR_ROW_MRS_ORDER_UNCONF, CLR_COL).Interior.Color = clrMrsOrderUnconf
    wsIn.Cells(CLR_ROW_MRS_ORDER_UNCONF, CLR_COL).Font.Color = GetContrastColor(clrMrsOrderUnconf)
    wsIn.Cells(CLR_ROW_HEADER, CLR_COL).Interior.Color = clrHeader
    wsIn.Cells(CLR_ROW_HEADER, CLR_COL).Font.Color = GetContrastColor(clrHeader)

    workerCount = CLng(Val(wsIn.Range("B9").Value))
    If workerCount < 1 Then workerCount = 1
    If workerCount > 10 Then workerCount = 10
    SyncWorkerIdInputs wsIn, workerCount

    opCount = CLng(Val(wsIn.Range("B8").Value))
    If opCount < 1 Then opCount = 1
    If opCount > 20 Then opCount = 20
    SyncOperationRows wsIn, opCount

    SyncPauseInputCell wsIn, wsHist, clrLocked, clrEditHasColor, clrEditable
End Sub

Private Sub RefreshResultSheetColors(ByVal wsOut As Worksheet, ByVal clrLocked As Long)
    wsOut.Cells.Interior.Color = clrLocked
    wsOut.Cells.Font.Color = GetContrastColor(clrLocked)
End Sub

Private Sub RefreshDisclaimerSheetColors(ByVal wsDisclaimer As Worksheet, ByVal clrLocked As Long)
    wsDisclaimer.Cells.Interior.Color = clrLocked
    wsDisclaimer.Cells.Font.Color = GetContrastColor(clrLocked)
End Sub

Private Sub RefreshHistorySheetColors(ByVal wsHist As Worksheet, ByVal clrLocked As Long, _
    ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long, ByVal clrMrsOrder As Long, ByVal clrMrsOrderUnconf As Long, ByVal clrHeader As Long)

    Dim lastRow As Long, r As Long, idx As Long
    Dim editCols As Variant

    lastRow = LastContentRow(wsHist, 1)
    wsHist.Cells.Interior.Color = clrLocked
    wsHist.Cells.Font.Color = GetContrastColor(clrLocked)

    wsHist.Rows("1:3").Interior.Color = clrHeader
    wsHist.Rows("1:3").Font.Color = GetContrastColor(clrHeader)

    Dim sConf As String, sUnconf As String
    sConf = UW(1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1086)
    sUnconf = UW(1053, 1045, 32, 1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1086)
    Dim rngConf As Range, rngUnconf As Range
    For r = 4 To lastRow
        Dim rowRng As Range
        Set rowRng = wsHist.Range(wsHist.Cells(r, 2), wsHist.Cells(r, 22))
        If Trim$(CStr(wsHist.Cells(r, 2).Value)) = sConf Then
            If rngConf Is Nothing Then Set rngConf = rowRng Else Set rngConf = Union(rngConf, rowRng)
        ElseIf Trim$(CStr(wsHist.Cells(r, 2).Value)) = sUnconf Then
            If rngUnconf Is Nothing Then Set rngUnconf = rowRng Else Set rngUnconf = Union(rngUnconf, rowRng)
        End If
    Next r
    If Not rngConf Is Nothing Then
        rngConf.Interior.Color = clrMrsOrder
        rngConf.Font.Color = GetContrastColor(clrMrsOrder)
    End If
    If Not rngUnconf Is Nothing Then
        rngUnconf.Interior.Color = clrMrsOrderUnconf
        rngUnconf.Font.Color = GetContrastColor(clrMrsOrderUnconf)
    End If

    editCols = Array(5, 7, 12, 16, 18, 19)
    For idx = LBound(editCols) To UBound(editCols)
        For r = 2 To lastRow
            If Not wsHist.Cells(r, editCols(idx)).Locked Then
                ApplyEditableStyle wsHist.Cells(r, editCols(idx)), clrEditHasColor, clrEditable
            End If
        Next r
    Next idx
End Sub

Private Sub RefreshMRSSheetColors(ByVal wsMRS As Worksheet, ByVal clrLocked As Long, _
    ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long, _
    ByVal clrMrsHeader As Long, ByVal clrMrsSub As Long, _
    ByVal clrMrsOrder As Long, ByVal clrMrsOrderUnconf As Long, ByVal clrHeader As Long)

    Dim lastRow As Long, r As Long, idx As Long
    Dim editCols As Variant

    lastRow = LastContentRow(wsMRS, 1)
    wsMRS.Cells.Interior.Color = clrLocked
    wsMRS.Cells.Font.Color = GetContrastColor(clrLocked)

    wsMRS.Rows("1:3").Interior.Color = clrHeader
    wsMRS.Rows("1:3").Font.Color = GetContrastColor(clrHeader)

    Dim sConf As String
    sConf = UW(1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1086)
    Dim rngHdr As Range, rngSub As Range, rngConf As Range, rngUnconf As Range
    For r = 4 To lastRow
        Dim rowRng As Range
        Set rowRng = wsMRS.Range(wsMRS.Cells(r, 2), wsMRS.Cells(r, 14))
        If wsMRS.Cells(r, 2).Font.Size = 18 Then
            If rngHdr Is Nothing Then Set rngHdr = rowRng Else Set rngHdr = Union(rngHdr, rowRng)
        ElseIf wsMRS.Cells(r, 2).Font.Size = 16 Then
            If rngSub Is Nothing Then Set rngSub = rowRng Else Set rngSub = Union(rngSub, rowRng)
        ElseIf wsMRS.Cells(r, 5).Font.Size = 15 Then
            If Trim$(CStr(wsMRS.Cells(r, 2).Value)) = sConf Then
                If rngConf Is Nothing Then Set rngConf = rowRng Else Set rngConf = Union(rngConf, rowRng)
            Else
                If rngUnconf Is Nothing Then Set rngUnconf = rowRng Else Set rngUnconf = Union(rngUnconf, rowRng)
            End If
        End If
    Next r
    If Not rngHdr Is Nothing Then
        rngHdr.Interior.Color = clrMrsHeader
        rngHdr.Font.Color = GetContrastColor(clrMrsHeader)
    End If
    If Not rngSub Is Nothing Then
        rngSub.Interior.Color = clrMrsSub
        rngSub.Font.Color = GetContrastColor(clrMrsSub)
    End If
    If Not rngConf Is Nothing Then
        rngConf.Interior.Color = clrMrsOrder
        rngConf.Font.Color = GetContrastColor(clrMrsOrder)
    End If
    If Not rngUnconf Is Nothing Then
        rngUnconf.Interior.Color = clrMrsOrderUnconf
        rngUnconf.Font.Color = GetContrastColor(clrMrsOrderUnconf)
    End If

    editCols = Array(7, 9, 13)
    For idx = LBound(editCols) To UBound(editCols)
        For r = 3 To lastRow
            If Not wsMRS.Cells(r, editCols(idx)).Locked Then
                ApplyEditableStyle wsMRS.Cells(r, editCols(idx)), clrEditHasColor, clrEditable
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

Private Function BuildValidationDateFormula(ByVal d As Date) As String
    BuildValidationDateFormula = CStr(CLng(d))
End Function

Private Function BuildValidationDecimalFormula(ByVal valueNum As Double) As String
    Dim decSep As String
    Dim txt As String

    decSep = Application.International(xlDecimalSeparator)
    txt = Trim$(CStr(valueNum))
    txt = Replace$(txt, ".", decSep)
    txt = Replace$(txt, ",", decSep)
    BuildValidationDecimalFormula = txt
End Function

Private Sub ApplyDateValidation(ByVal target As Range, ByVal minDate As Date, ByVal maxDate As Date)
    With target.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:=CLng(minDate), Formula2:=CLng(maxDate)
        .IgnoreBlank = True
    End With
End Sub

Private Sub ApplyTimeValidation(ByVal target As Range)
    With target.Validation
        .Delete
        .Add Type:=xlValidateTime, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:=0, Formula2:=86399# / 86400#
        .IgnoreBlank = True
    End With
End Sub

Private Sub ApplyDecimalValidation(ByVal target As Range, ByVal minVal As Double, ByVal maxVal As Double)
    With target.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:=minVal, Formula2:=maxVal
        .IgnoreBlank = True
    End With
End Sub

Private Sub ApplyWholeValidation(ByVal target As Range, ByVal minVal As Long, ByVal maxVal As Long)
    With target.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:=minVal, Formula2:=maxVal
        .IgnoreBlank = True
    End With
End Sub

Public Sub RefreshWorkbookColors()
    On Error GoTo EH

    Dim wsDisclaimer As Worksheet, wsIn As Worksheet, wsOut As Worksheet, wsHist As Worksheet, wsMRS As Worksheet
    Set wsDisclaimer = ThisWorkbook.Worksheets(1)
    Set wsIn = ThisWorkbook.Worksheets(2)
    Set wsOut = ThisWorkbook.Worksheets(3)
    Set wsHist = ThisWorkbook.Worksheets(4)
    Set wsMRS = ThisWorkbook.Worksheets(5)

    Dim clrLocked As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrMrsOrderUnconf As Long
    Dim clrHeader As Long

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    wsDisclaimer.Unprotect UW(49, 49, 52, 55, 48, 57)
    wsIn.Unprotect UW(49, 49, 52, 55, 48, 57)
    wsOut.Unprotect UW(49, 49, 52, 55, 48, 57)
    wsHist.Unprotect UW(49, 49, 52, 55, 48, 57)
    wsMRS.Unprotect UW(49, 49, 52, 55, 48, 57)

    EnsureColorSettings wsIn
    ReadAllColors wsIn, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrMrsOrderUnconf, clrHeader

    RefreshDisclaimerSheetColors wsDisclaimer, clrLocked
    RefreshInputSheetColors wsIn, wsHist, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrMrsOrderUnconf, clrHeader
    RefreshResultSheetColors wsOut, clrLocked
    RefreshHistorySheetColors wsHist, clrLocked, clrEditHasColor, clrEditable, clrMrsOrder, clrMrsOrderUnconf, clrHeader
    RefreshMRSSheetColors wsMRS, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrMrsOrderUnconf, clrHeader

Cleanup:
    On Error Resume Next
    wsDisclaimer.Protect UW(49, 49, 52, 55, 48, 57)
    wsIn.Protect UW(49, 49, 52, 55, 48, 57)
    wsOut.Protect UW(49, 49, 52, 55, 48, 57)
    wsHist.Protect UW(49, 49, 52, 55, 48, 57)
    wsMRS.Protect UW(49, 49, 52, 55, 48, 57)
    Application.ScreenUpdating = True
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

Public Sub PickColorMrsOrderUnconf()
    PickCellColor CLR_ROW_MRS_ORDER_UNCONF
End Sub

Public Sub PickColorHeader()
    PickCellColor CLR_ROW_HEADER
End Sub

Public Function UW(ParamArray codes() As Variant) As String
    Dim i As Long
    For i = LBound(codes) To UBound(codes)
        UW = UW & ChrW(CLng(codes(i)))
    Next i
End Function

Private Function ZQ() As Boolean
    ZQ = False
    On Error Resume Next
    Dim v As String
    v = Trim$(CStr(ThisWorkbook.Worksheets(2).Range(UW(65, 49)).Value))
    If Len(v) < 2 Then Exit Function
    Dim ec As Variant
    ec = Array(181, 239, 280, 309, 329, 208, 218, 268, 306, 353, 198, 179, 257, 274, 310, 129, _
               232, 292, 240, 1288, 1169, 1217, 1251, 1292, 1324, 1200, 1219, 1257, 1282, _
               277, 1140, 180, 1227, 254)
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
    Set wsIn = ThisWorkbook.Worksheets(2)
    Set wsOut = ThisWorkbook.Worksheets(3)
    Set wsHist = ThisWorkbook.Worksheets(4)

    wsIn.Unprotect UW(49, 49, 52, 55, 48, 57)
    wsOut.Unprotect UW(49, 49, 52, 55, 48, 57)
    wsHist.Unprotect UW(49, 49, 52, 55, 48, 57)
    Application.EnableEvents = False

    EnsureColorSettings wsIn
    Dim clrLocked As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrMrsOrderUnconf As Long
    Dim clrHeader As Long
    ReadAllColors wsIn, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrMrsOrderUnconf, clrHeader

    ClearResultArea wsOut
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
    lunchDurMin = Val(Replace(CStr(wsIn.Range("B12").Value), ",", "."))
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

        Dim durVal As Double, durUnit As String, timeMode As String
        durVal = Val(Replace(CStr(wsIn.Cells(opRow, 11).Value), ",", "."))
        If durVal < 0 Then durVal = 0
        durUnit = NormalizeUnit(wsIn.Cells(opRow, 12).Value)
        timeMode = LCase$(Trim$(CStr(wsIn.Cells(opRow, 13).Value)))

        Dim breakVal As Double, breakUnit As String
        breakVal = Val(Replace(CStr(wsIn.Cells(opRow, 14).Value), ",", "."))
        If breakVal < 0 Then breakVal = 0
        breakUnit = NormalizeUnit(wsIn.Cells(opRow, 15).Value)

        Dim workerSpec As String
        workerSpec = Trim$(CStr(wsIn.Cells(opRow, 16).Value))

        Dim displayDurVal As Double
        displayDurVal = durVal
        Dim isPerWorker As Boolean
        isPerWorker = False
        If InStr(timeMode, LCase$(UW(1085, 1072, 32, 1082, 1072, 1078, 1076))) > 0 Then
            isPerWorker = True
        End If

        If Not isPerWorker And workerCount > 1 Then
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

        Dim breakMin As Double
        If breakUnit = "hour" Then
            breakMin = breakVal * 60
        Else
            breakMin = breakVal
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
                    wsOut.Cells(outRow, 5).Value = breakMin
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

    wsOut.Range("E2:E" & outRow - 1).NumberFormat = "0.00"
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

    AppendResultToHistory wsHist, wsOut, wsIn, outRow - 1, zRow + 5, workerCount, lunch1, lunch2, lunchDurMin, primaryIsMin, clrEditHasColor, clrEditable, clrMrsOrder, clrMrsOrderUnconf
    RefreshResultSheetColors wsOut, clrLocked
    RefreshHistorySheetColors wsHist, clrLocked, clrEditHasColor, clrEditable, clrMrsOrder, clrMrsOrderUnconf, clrHeader

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
    wsIn.Cells(4, 14).Locked = False
    ApplyEditableStyle wsIn.Cells(4, 14), clrEditHasColor, clrEditable
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
    wsOut.Cells(1, 5).Value = UW(1055, 1072, 1091, 1079, 1072, 32, 40, 1084, 1080, 1085, 41)
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
    SetRangeBoldSafe wsOut.Range("A1:V1"), True, "EnsureHeaders wsOut"

    wsHist.Cells(3, 2).Value = UW(1048, 1089, 1090, 1086, 1088, 1080, 1103)
    SetCellBoldSafe wsHist.Cells(3, 2), True, "EnsureHeaders wsHist"
End Sub

Public Sub ClearResultAndHistory()
    If Not ZQ() Then
        MsgBox UW(1054, 1096, 1080, 1073, 1082, 1072, 33), vbCritical
        Exit Sub
    End If
    On Error GoTo EH
    Dim wsIn As Worksheet, wsR As Worksheet, wsH As Worksheet
    Set wsIn = ThisWorkbook.Worksheets(2)
    Set wsR = ThisWorkbook.Worksheets(3)
    Set wsH = ThisWorkbook.Worksheets(4)
    wsIn.Unprotect UW(49, 49, 52, 55, 48, 57)
    EnsureColorSettings wsIn
    wsIn.Range("B5").Value = Date
    wsIn.Range("B6").Value = TimeSerial(8, 0, 0)
    wsIn.Range("B7").Value = Date
    Dim clrLocked As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrMrsOrderUnconf As Long
    Dim clrHeader As Long
    ReadAllColors wsIn, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrMrsOrderUnconf, clrHeader
    wsIn.Protect UW(49, 49, 52, 55, 48, 57)
    wsR.Unprotect UW(49, 49, 52, 55, 48, 57)
    ClearResultArea wsR
    wsR.Protect UW(49, 49, 52, 55, 48, 57)
    wsH.Unprotect UW(49, 49, 52, 55, 48, 57)
    ClearHistoryArea wsH, clrLocked
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

Private Sub ClearResultArea(ByVal wsOut As Worksheet)
    Dim lastRow As Long
    lastRow = wsOut.UsedRange.Row + wsOut.UsedRange.Rows.Count - 1
    If lastRow >= 2 Then
        Dim prevEvents As Boolean
        prevEvents = Application.EnableEvents
        Application.EnableEvents = False
        wsOut.Rows("2:" & CStr(lastRow)).Delete Shift:=xlUp
        Application.EnableEvents = prevEvents
    End If
End Sub

Private Sub ClearHistoryArea(ByVal wsHist As Worksheet, ByVal clrLocked As Long)
    Dim lastRow As Long
    lastRow = wsHist.UsedRange.Row + wsHist.UsedRange.Rows.Count - 1
    If lastRow >= 4 Then
        Dim prevEvents As Boolean
        prevEvents = Application.EnableEvents
        Application.EnableEvents = False
        wsHist.Rows("4:" & CStr(lastRow)).Delete Shift:=xlUp
        Application.EnableEvents = prevEvents
    End If
End Sub

Public Sub SyncWorkerIdInputs(ByVal wsIn As Worksheet, ByVal workerCount As Long)
    Dim i As Long

    If workerCount < 1 Then workerCount = 1
    If workerCount > 10 Then workerCount = 10

    EnsureColorSettings wsIn
    Dim clrLocked As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrMrsOrderUnconf As Long
    Dim clrHeader As Long
    ReadAllColors wsIn, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrMrsOrderUnconf, clrHeader

    wsIn.Range("D3").Value = UW(1048, 1089, 1087, 1086, 1083, 1085, 1080, 1090, 1077, 1083, 1100)
    SetRangeBoldSafe wsIn.Range("D3:E3"), True, "SyncWorkerIdInputs"
    wsIn.Range("E4:E13").NumberFormat = "@"

    For i = 1 To 10
        Dim rowNum As Long
        rowNum = 3 + i

        wsIn.Cells(rowNum, 4).Value = UW(1048, 1089, 1087, 1086, 1083, 1085, 1080, 1090, 1077, 1083, 1100, 32) & i
        StyleWorkerInputRow wsIn, rowNum, (i <= workerCount), clrLocked, clrEditHasColor, clrEditable
    Next i
End Sub

Public Sub SyncOperationRows(ByVal wsIn As Worksheet, ByVal opCount As Long)
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Cleanup

    If opCount < 1 Then opCount = 1
    If opCount > 20 Then opCount = 20

    EnsureColorSettings wsIn
    Dim clrLocked As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrMrsOrderUnconf As Long
    Dim clrHeader As Long
    ReadAllColors wsIn, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrMrsOrderUnconf, clrHeader

    Dim firstDurUnit As String, firstType As String, firstBreakUnit As String
    firstDurUnit = Trim$(CStr(wsIn.Cells(4, 12).Value))
    firstType = Trim$(CStr(wsIn.Cells(4, 13).Value))
    firstBreakUnit = Trim$(CStr(wsIn.Cells(4, 15).Value))

    Dim i As Long, r As Long, c As Long
    Dim editCols As Variant
    editCols = Array(8, 9, 10, 11, 14, 16)
    Dim syncCols As Variant
    syncCols = Array(12, 13, 15)

    For i = 1 To opCount
        r = i + 3
        wsIn.Cells(r, 7).Value = i
        wsIn.Cells(r, 7).Font.Color = GetContrastColor(clrLocked)
        If Trim$(CStr(wsIn.Cells(r, 9).Value)) = "" Then
            wsIn.Cells(r, 9).Value = UW(1054, 1087, 1077, 1088, 1072, 1094, 1080, 1103) & " " & i
        End If
        If Trim$(CStr(wsIn.Cells(r, 11).Value)) = "" Then
            wsIn.Cells(r, 11).Value = 0
        End If
        If i > 1 Then
            If firstDurUnit <> "" Then wsIn.Cells(r, 12).Value = firstDurUnit
            If firstType <> "" Then wsIn.Cells(r, 13).Value = firstType
            wsIn.Cells(r, 14).Value = 0
            If firstBreakUnit <> "" Then wsIn.Cells(r, 15).Value = firstBreakUnit
            wsIn.Cells(r, 16).Value = ""
        End If
        For c = LBound(editCols) To UBound(editCols)
            wsIn.Cells(r, editCols(c)).Locked = False
            ApplyEditableStyle wsIn.Cells(r, editCols(c)), clrEditHasColor, clrEditable
            wsIn.Cells(r, editCols(c)).Borders.LineStyle = xlContinuous
        Next c
        If i = 1 Then
            Dim wsHist As Worksheet
            Set wsHist = ThisWorkbook.Worksheets(4)
            Dim histLastRow As Long
            histLastRow = wsHist.Cells(wsHist.Rows.Count, 2).End(xlUp).Row
            If histLastRow <= 3 Then
                wsIn.Cells(r, 14).Locked = True
                wsIn.Cells(r, 14).Interior.Color = clrLocked
            End If
        End If

        For c = LBound(syncCols) To UBound(syncCols)
            wsIn.Cells(r, syncCols(c)).Borders.LineStyle = xlContinuous
            wsIn.Cells(r, syncCols(c)).Font.Color = GetContrastColor(clrLocked)
            If i = 1 Then
                If clrEditHasColor Then
                    wsIn.Cells(r, syncCols(c)).Interior.Color = clrEditable
                    wsIn.Cells(r, syncCols(c)).Font.Color = GetContrastColor(clrEditable)
                Else
                    wsIn.Cells(r, syncCols(c)).Interior.Pattern = xlNone
                    wsIn.Cells(r, syncCols(c)).Font.Color = 0
                End If
                wsIn.Cells(r, syncCols(c)).Locked = False
            Else
                wsIn.Cells(r, syncCols(c)).Interior.Color = clrLocked
                wsIn.Cells(r, syncCols(c)).Locked = True
            End If
        Next c
    Next i

    If opCount + 4 <= 23 Then
        Dim unusedRange As Range
        Set unusedRange = wsIn.Range(wsIn.Cells(opCount + 4, 7), wsIn.Cells(23, 16))
        unusedRange.ClearContents
        ApplyLockedStyle unusedRange, clrLocked
        unusedRange.Borders.LineStyle = xlNone
        unusedRange.Locked = True
    End If

Cleanup:
    Application.EnableEvents = prevEvents
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
    ByVal clrLocked As Long, _
    ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long)
    With wsIn.Cells(rowNum, 4)
        If isVisible Then
            .Font.Color = GetContrastColor(clrLocked)
        Else
            .Font.Color = clrLocked
            .ClearContents
        End If
        .Locked = True
    End With

    With wsIn.Cells(rowNum, 5)
        .NumberFormat = "00000000"
        If isVisible Then
            ApplyEditableStyle wsIn.Cells(rowNum, 5), clrEditHasColor, clrEditable
            .Borders.LineStyle = xlContinuous
            .Locked = False
        Else
            ApplyLockedStyle wsIn.Cells(rowNum, 5), clrLocked
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

    For daySerial = CLng(Int(opStart)) - 1 To CLng(Int(opEnd)) + 1
        lStart = daySerial + TimePart(lunch1)
        lEnd = lStart + lunchDurDays

        If opStart < lStart And opEnd > lStart Then
            opEnd = opEnd + lunchDurDays
            crossedLunch = True
            Exit For
        End If
    Next daySerial

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
    For daySerial = CLng(Int(st)) - 1 To CLng(Int(st)) + 1
        Dim lStart As Date, lEnd As Date
        lStart = daySerial + TimePart(lunch1)
        lEnd = lStart + lunchDurDays
        If st >= lStart - 0.00001 And st < lEnd Then
            st = lEnd
            Exit For
        End If
    Next daySerial

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
        ByVal clrMrsOrder As Long, _
        ByVal clrMrsOrderUnconf As Long)

    Dim NextRow As Long
    NextRow = wsHist.Cells(wsHist.Rows.Count, 2).End(xlUp).Row + 1
    If NextRow < 4 Then NextRow = 4

    If NextRow > 4 Then NextRow = NextRow + 1

    Dim orderNum As String, orderName As String
    orderNum = Trim$(CStr(wsIn.Range("B3").Value))
    orderName = Trim$(CStr(wsIn.Range("B4").Value))
    
    wsHist.Range(wsHist.Cells(NextRow, 2), wsHist.Cells(NextRow, 4)).Merge
    wsHist.Cells(NextRow, 2).Value = UW(1053, 1045, 32, 1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1086)
    With wsHist.Cells(NextRow, 2).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=UW(1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1086) & "," & UW(1053, 1045, 32, 1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1086)
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    wsHist.Range(wsHist.Cells(NextRow, 2), wsHist.Cells(NextRow, 4)).Locked = False

    wsHist.Cells(NextRow, 5).Value = UW(1047, 1072, 1082, 1072, 1079) & ":"
    
    wsHist.Range(wsHist.Cells(NextRow, 6), wsHist.Cells(NextRow, 7)).Merge
    wsHist.Range(wsHist.Cells(NextRow, 6), wsHist.Cells(NextRow, 7)).NumberFormat = "0"
    wsHist.Cells(NextRow, 6).Value = orderNum
    
    wsHist.Range(wsHist.Cells(NextRow, 8), wsHist.Cells(NextRow, 22)).Merge
    wsHist.Cells(NextRow, 8).Value = UW(1053, 1072, 1080, 1084, 1077, 1085, 1086, 1074, 1072, 1085, 1080, 1077) & ": " & orderName & _
        "  |  " & UW(1057, 1092, 1086, 1088, 1084, 1080, 1088, 1086, 1074, 1072, 1085) & ": " & Format$(Now, "dd.mm.yyyy hh:nn:ss")

    SetCellBoldSafe wsHist.Cells(NextRow, 5), True, "AppendHistory Label"
    SetCellBoldSafe wsHist.Cells(NextRow, 6), True, "AppendHistory OrderNum"
    SetCellBoldSafe wsHist.Cells(NextRow, 8), True, "AppendHistory Title"
    wsHist.Range(wsHist.Cells(NextRow, 2), wsHist.Cells(NextRow, 22)).Borders.LineStyle = xlContinuous
    wsHist.Range(wsHist.Cells(NextRow, 2), wsHist.Cells(NextRow, 22)).HorizontalAlignment = -4108
    wsHist.Range(wsHist.Cells(NextRow, 2), wsHist.Cells(NextRow, 22)).Interior.Color = clrMrsOrderUnconf
    wsHist.Range(wsHist.Cells(NextRow, 2), wsHist.Cells(NextRow, 22)).Font.Color = GetContrastColor(clrMrsOrderUnconf)
    NextRow = NextRow + 1

    wsOut.Range("A1:V1").Copy Destination:=wsHist.Cells(NextRow, 1)
    NextRow = NextRow + 1
    Dim dataStartRow As Long, dataEndRow As Long
    dataStartRow = NextRow
    dataEndRow = NextRow + (lastDataRow - 2)

    wsOut.Range("A2:V" & lastDataRow).Copy Destination:=wsHist.Cells(dataStartRow, 1)

    wsHist.Cells(dataStartRow, 15).Value = wsOut.Cells(2, 15).Value
    If dataEndRow > dataStartRow Then
        wsHist.Range("O" & (dataStartRow + 1) & ":O" & dataEndRow).Formula = "=$O$" & dataStartRow
    End If

    Dim prevDataEndRow As Long, scanRow As Long
    prevDataEndRow = 0
    For scanRow = dataStartRow - 1 To 4 Step -1
        If IsNumeric(wsHist.Cells(scanRow, 2).Value) Then
            If Len(Trim$(CStr(wsHist.Cells(scanRow, 20).Value))) > 0 And Len(Trim$(CStr(wsHist.Cells(scanRow, 21).Value))) > 0 Then
                prevDataEndRow = scanRow
                Exit For
            End If
        End If
    Next scanRow

    If prevDataEndRow > 0 Then
        Dim rawDTFirst As String
        rawDTFirst = "(T" & prevDataEndRow & "+U" & prevDataEndRow & "+E" & dataStartRow & "/1440)"
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
            rawDT = "(T" & (rr - 1) & "+U" & (rr - 1) & "+E" & rr & "/1440)"
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

    wsHist.Range("E" & dataStartRow & ":E" & dataEndRow).NumberFormat = "0.00"
    wsHist.Range("O" & dataStartRow & ":O" & dataEndRow).NumberFormat = "dd"".""mm"".""yyyy"
    wsHist.Range("R" & dataStartRow & ":R" & dataEndRow).NumberFormat = "dd"".""mm"".""yyyy"
    wsHist.Range("S" & dataStartRow & ":S" & dataEndRow).NumberFormat = "h:mm:ss"
    wsHist.Range("T" & dataStartRow & ":T" & dataEndRow).NumberFormat = "dd"".""mm"".""yyyy"
    wsHist.Range("U" & dataStartRow & ":U" & dataEndRow).NumberFormat = "h:mm:ss"

    Dim headerRow As Long
    headerRow = dataStartRow - 1
    wsHist.Range(wsHist.Cells(headerRow, 2), wsHist.Cells(dataEndRow, 22)).Borders.LineStyle = xlContinuous

    wsHist.Range(wsHist.Cells(headerRow, 7), wsHist.Cells(dataEndRow, 7)).Borders(7).LineStyle = xlContinuous
    wsHist.Range(wsHist.Cells(headerRow, 7), wsHist.Cells(dataEndRow, 7)).Borders(7).Weight = 4
    wsHist.Range(wsHist.Cells(headerRow, 21), wsHist.Cells(dataEndRow, 21)).Borders(10).LineStyle = xlContinuous
    wsHist.Range(wsHist.Cells(headerRow, 21), wsHist.Cells(dataEndRow, 21)).Borders(10).Weight = 4

    NextRow = NextRow + (lastDataRow - 1) + 1

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

    Dim editCol As Variant, ec As Long
    editCol = Array(5, 7, 12, 16)
    For ec = LBound(editCol) To UBound(editCol)
        For rr = dataStartRow To dataEndRow
            If editCol(ec) = 5 And rr = dataStartRow And prevDataEndRow = 0 Then GoTo SkipCell
            If Not wsHist.Cells(rr, editCol(ec)).HasFormula Then
                wsHist.Cells(rr, editCol(ec)).Locked = False
                If clrEditHasColor Then
                    wsHist.Cells(rr, editCol(ec)).Interior.Color = clrEditable
                    wsHist.Cells(rr, editCol(ec)).Font.Color = GetContrastColor(clrEditable)
                Else
                    wsHist.Cells(rr, editCol(ec)).Interior.Pattern = xlNone
                    wsHist.Cells(rr, editCol(ec)).Font.Color = 0
                End If
                Select Case editCol(ec)
                    Case 5, 12
                        wsHist.Cells(rr, editCol(ec)).NumberFormat = "0.00"
                        ApplyDecimalValidation wsHist.Cells(rr, editCol(ec)), 0, 999999.99
                    Case 7, 16
                        ApplyWholeValidation wsHist.Cells(rr, editCol(ec)), 0, 99999999
                End Select
            End If
SkipCell:
        Next rr
    Next ec

    wsHist.Range(wsHist.Cells(dataStartRow, 12), wsHist.Cells(dataEndRow, 12)).NumberFormat = "0.00"

    If prevDataEndRow = 0 Then
        wsHist.Cells(dataStartRow, 18).Locked = False
        If clrEditHasColor Then
            wsHist.Cells(dataStartRow, 18).Interior.Color = clrEditable
            wsHist.Cells(dataStartRow, 18).Font.Color = GetContrastColor(clrEditable)
        Else
            wsHist.Cells(dataStartRow, 18).Interior.Pattern = xlNone
            wsHist.Cells(dataStartRow, 18).Font.Color = 0
        End If
        ApplyDateValidation wsHist.Cells(dataStartRow, 18), DateSerial(2000, 1, 1), DateSerial(2099, 12, 31)
        wsHist.Cells(dataStartRow, 19).Locked = False
        If clrEditHasColor Then
            wsHist.Cells(dataStartRow, 19).Interior.Color = clrEditable
            wsHist.Cells(dataStartRow, 19).Font.Color = GetContrastColor(clrEditable)
        Else
            wsHist.Cells(dataStartRow, 19).Interior.Pattern = xlNone
            wsHist.Cells(dataStartRow, 19).Font.Color = 0
        End If
        ApplyTimeValidation wsHist.Cells(dataStartRow, 19)
    End If

    Dim blockStart As Long
    blockStart = headerRow - 1
    wsHist.Range(wsHist.Cells(blockStart, 1), wsHist.Cells(z7Start + 5, 22)).Font.Size = 14

    wsHist.Columns("A:A").ColumnWidth = 2.7
    wsHist.Columns("B:B").ColumnWidth = 3.5
    wsHist.Columns("C:C").ColumnWidth = 60
    wsHist.Columns("D:D").ColumnWidth = 8
    wsHist.Columns("E:E").ColumnWidth = 14
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

    Application.ScreenUpdating = True
    ClearMRS
    DoEvents

    Dim wsMRS As Worksheet
    Set wsMRS = ThisWorkbook.Worksheets(5)

    Dim filePath As Variant
    If MsgBox(UW(1053, 1077, 32, 1079, 1072, 1075, 1088, 1091, 1078, 1072, 1081, 1090, 1077, 32, 1089, 1083, 1080, 1096, 1082, 1086, 1084, 32, 1084, 1085, 1086, 1075, 1086, 32, 1076, 1072, 1085, 1085, 1099, 1093, 33), vbOKCancel + vbExclamation) = vbCancel Then Exit Sub

    filePath = Application.GetOpenFilename(FileFilter:=UW(1060, 1072, 1081, 1083, 1099) & " Excel (*.xlsx), *.xlsx")
    If filePath = False Then Exit Sub

    Dim wsIn As Worksheet
    Set wsIn = ThisWorkbook.Worksheets(2)
    stage = "Ensure color settings"
    wsIn.Unprotect UW(49, 49, 52, 55, 48, 57)
    EnsureColorSettings wsIn
    Dim clrLocked As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrMrsOrderUnconf As Long
    Dim clrHeader As Long
    stage = "Read colors from settings"
    ReadAllColors wsIn, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrMrsOrderUnconf, clrHeader
    
    Dim lunch1 As Date, lunch2 As Date, hasLunch2 As Boolean, lunchDurDays As Double
    Dim lunchDurMin As Double
    lunchDurMin = Val(Replace(CStr(wsIn.Range("B12").Value), ",", "."))
    ParseLunchParams wsIn.Range("B10").Value, wsIn.Range("B11").Value, lunchDurMin, lunch1, lunch2, hasLunch2, lunchDurDays

   wsIn.Protect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim calcState As Long
    calcState = Application.Calculation
    Application.Calculation = xlCalculationManual
    wsMRS.Unprotect UW(49, 49, 52, 55, 48, 57)
    ThisWorkbook.Unprotect UW(49, 49, 52, 55, 48, 57)

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
        If i Mod 1000 = 0 Then
            Application.StatusBar = UW(1055, 1072, 1088, 1089, 1080, 1085, 1075, 32, 77, 82, 83, 58, 32, 1063, 1090, 1077, 1085, 1080, 1077, 32, 1076, 1072, 1085, 1085, 1099, 1093, 32) & Int((i / totalRows) * 100) & "%"
            DoEvents
        End If
        If IsError(data(i + 1, 5)) Then GoTo NextRow
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
        If Not IsError(data(i + 1, 3)) Then arrWName(validCount) = CStr(data(i + 1, 3)) Else arrWName(validCount) = ""
        If Not IsError(data(i + 1, 4)) Then arrWID(validCount) = CStr(data(i + 1, 4)) Else arrWID(validCount) = ""
        If Not IsEmpty(data(i + 1, 8)) And Not IsError(data(i + 1, 8)) Then
            If IsDate(data(i + 1, 8)) Then
                arrSDate(validCount) = CDate(data(i + 1, 8))
            ElseIf IsNumeric(data(i + 1, 8)) Then
                arrSDate(validCount) = CDate(CDbl(data(i + 1, 8)))
            End If
        End If
        If Not IsEmpty(data(i + 1, 9)) And Not IsError(data(i + 1, 9)) Then
            If IsDate(data(i + 1, 9)) Then
                arrSTime(validCount) = CDate(data(i + 1, 9))
            ElseIf IsNumeric(data(i + 1, 9)) Then
                arrSTime(validCount) = CDate(CDbl(data(i + 1, 9)))
            End If
        End If
        If Not IsEmpty(data(i + 1, 10)) And Not IsError(data(i + 1, 10)) Then
            If IsDate(data(i + 1, 10)) Then
                arrEDate(validCount) = CDate(data(i + 1, 10))
            ElseIf IsNumeric(data(i + 1, 10)) Then
                arrEDate(validCount) = CDate(CDbl(data(i + 1, 10)))
            End If
        End If
        If Not IsEmpty(data(i + 1, 11)) And Not IsError(data(i + 1, 11)) Then
            If IsDate(data(i + 1, 11)) Then
                arrETime(validCount) = CDate(data(i + 1, 11))
            ElseIf IsNumeric(data(i + 1, 11)) Then
                arrETime(validCount) = CDate(CDbl(data(i + 1, 11)))
            End If
        End If
        If Not IsEmpty(data(i + 1, 12)) And Not IsError(data(i + 1, 12)) Then
            If IsNumeric(data(i + 1, 12)) Then arrDur(validCount) = CDbl(data(i + 1, 12))
        End If
        If Not IsError(data(i + 1, 14)) Then arrWPlace(validCount) = CStr(data(i + 1, 14)) Else arrWPlace(validCount) = ""
        If Not IsError(data(i + 1, 16)) Then arrPDTV(validCount) = CStr(data(i + 1, 16)) Else arrPDTV(validCount) = ""
NextRow:
    Next i

    totalRows = validCount
    If totalRows = 0 Then
        MsgBox UW(1053, 1077, 1090, 32, 1079, 1072, 1082, 1072, 1079, 1086, 1074, 32, 1089, 32, 1095, 1080, 1089, 1083, 1086, 1074, 1099, 1084, 32, 1085, 1086, 1084, 1077, 1088, 1086, 1084), vbInformation
        GoTo Done
    End If

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
    If dateCount > 1 Then QuickSortLong dateKeys, 1, dateCount

    stage = "Prompt for dates"
    If dateCount > 1 Then
        Dim dlgDate As Object
        Set dlgDate = ThisWorkbook.DialogSheets.Add
        dlgDate.DialogFrame.Width = 1000
        dlgDate.DialogFrame.Height = 500
        dlgDate.DialogFrame.Characters.Text = ""
        
        Dim lblDateDesc As Object
        Set lblDateDesc = dlgDate.Labels.Add(dlgDate.DialogFrame.Left + 15, dlgDate.DialogFrame.Top + 15, dlgDate.DialogFrame.Width - 120, 35)
        lblDateDesc.Text = UW(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1076, 1085, 1080, 32, 1076, 1083, 1103, 32, 1079, 1072, 1075, 1088, 1091, 1079, 1082, 1080, 58, 10, 40, 1050, 1083, 1080, 1082, 1085, 1080, 1090, 1077, 32, 1087, 1086, 32, 1089, 1090, 1088, 1086, 1082, 1072, 1084, 32, 1076, 1083, 1103, 32, 1074, 1099, 1073, 1086, 1088, 1072, 32, 1085, 1077, 1089, 1082, 1086, 1083, 1100, 1082, 1080, 1093, 41)
        
        Dim dBtnDate As Object
        Dim btnDateTop As Double
        btnDateTop = dlgDate.DialogFrame.Top + 20
        For Each dBtnDate In dlgDate.Buttons
            dBtnDate.Left = dlgDate.DialogFrame.Left + dlgDate.DialogFrame.Width - dBtnDate.Width - 15
            dBtnDate.Top = btnDateTop
            btnDateTop = btnDateTop + dBtnDate.Height + 10
        Next dBtnDate
        
        Dim lbDate As Object
        Set lbDate = dlgDate.ListBoxes.Add(dlgDate.DialogFrame.Left + 15, dlgDate.DialogFrame.Top + 55, dlgDate.DialogFrame.Width - 120, dlgDate.DialogFrame.Height - 100)
        lbDate.MultiSelect = xlSimple
        
        Dim btnDateSelAll As Object
        Set btnDateSelAll = dlgDate.Buttons.Add(dlgDate.DialogFrame.Left + 15, dlgDate.DialogFrame.Top + dlgDate.DialogFrame.Height - 35, 120, 22)
        btnDateSelAll.Caption = UW(1042, 1099, 1073, 1088, 1072, 1090, 1100, 32, 1074, 1089, 1077)
        btnDateSelAll.OnAction = "SelectAllBrigades"
        
        Dim btnDateClearAll As Object
        Set btnDateClearAll = dlgDate.Buttons.Add(dlgDate.DialogFrame.Left + 145, dlgDate.DialogFrame.Top + dlgDate.DialogFrame.Height - 35, 120, 22)
        btnDateClearAll.Caption = UW(1054, 1095, 1080, 1089, 1090, 1080, 1090, 1100, 32, 1074, 1099, 1073, 1086, 1088)
        btnDateClearAll.OnAction = "ClearAllBrigades"
        
        Dim dIdx As Long
        For dIdx = 1 To dateCount
            lbDate.AddItem Format(CDate(CDbl(dateKeys(dIdx))), "dd"".""mm"".""yyyy")
        Next dIdx
        
        Application.ScreenUpdating = True
        Dim dlgDateResult As Boolean
        dlgDateResult = dlgDate.Show
        Application.ScreenUpdating = False
        
        Dim anyDateSel As Boolean: anyDateSel = False
        Dim newDateCount As Long: newDateCount = 0
        Dim newDateKeys() As Long
        ReDim newDateKeys(1 To dateCount)
        
        If dlgDateResult Then
            For dIdx = 1 To dateCount
                If lbDate.Selected(dIdx) Then
                    anyDateSel = True
                    newDateCount = newDateCount + 1
                    newDateKeys(newDateCount) = dateKeys(dIdx)
                End If
            Next dIdx
        Else
            Application.DisplayAlerts = False
            dlgDate.Delete
            Application.DisplayAlerts = True
            GoTo Done
        End If
        
        If Not anyDateSel Then
            Application.DisplayAlerts = False
            dlgDate.Delete
            Application.DisplayAlerts = True
            GoTo Done
        End If
        
        Application.DisplayAlerts = False
        dlgDate.Delete
        Application.DisplayAlerts = True
        
        dateCount = newDateCount
        dateKeys = newDateKeys
    End If

    Dim outRow As Long
    outRow = 3

    Dim d As Long
    For d = 1 To dateCount
        Dim curDateKey As Long
        curDateKey = dateKeys(d)

        Application.StatusBar = UW(1055, 1072, 1088, 1089, 1080, 1085, 1075, 32, 77, 82, 83, 58, 32, 1060, 1086, 1088, 1084, 1080, 1088, 1086, 1074, 1072, 1085, 1080, 1077, 32, 1074, 1099, 1074, 1086, 1076, 1072, 32) & Int((d / dateCount) * 100) & "%"
        DoEvents

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

    stage = "Build brigade keys for date " & CStr(d)
    Dim dictOrdBrigKey As Object
    Set dictOrdBrigKey = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    For Each key In dictOrderW.Keys
        Dim wIds As Variant
        wIds = dictOrderW(key).Keys
        Dim wCount As Long
        wCount = dictOrderW(key).Count
        If wCount > 1 Then QuickSortVariantStrings wIds, 0, wCount - 1

        Dim bKey As String: bKey = ""
        For i = 0 To wCount - 1
            If i > 0 Then bKey = bKey & ","
            bKey = bKey & CStr(wIds(i))
        Next i
        dictOrdBrigKey.Add CStr(key), bKey
    Next key

    stage = "Group brigades for date " & CStr(d)
    Dim dictBrigades As Object
    Set dictBrigades = CreateObject("Scripting.Dictionary")
    Dim dictBrigTime As Object
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

    If brigCount > 1 Then QuickSortDoubleString brigTimes, brigKeys, 1, brigCount

    stage = "Prompt for brigades date " & CStr(d)
    Dim brigNames() As String
    ReDim brigNames(1 To brigCount)
    Dim selBrig() As Boolean
    ReDim selBrig(1 To brigCount)
    
    Dim bDisp As Long
    For bDisp = 1 To brigCount
        Dim tmpBrigKey As String
        tmpBrigKey = brigKeys(bDisp)
        Dim tmpWorkerIds As Variant
        tmpWorkerIds = Split(tmpBrigKey, ",")
        Dim tmpFirstOrd As Variant
        For Each tmpFirstOrd In dictBrigades(tmpBrigKey).Keys: Exit For: Next
        Dim tmpWStr As String: tmpWStr = ""
        Dim twi As Long
        For twi = 0 To UBound(tmpWorkerIds)
            If twi > 0 Then tmpWStr = tmpWStr & ", "
            Dim tmpWName As String: tmpWName = ""
            If dictOrderW(CStr(tmpFirstOrd)).Exists(CStr(tmpWorkerIds(twi))) Then
                tmpWName = dictOrderW(CStr(tmpFirstOrd))(CStr(tmpWorkerIds(twi)))
            End If
            tmpWStr = tmpWStr & tmpWName & " (" & tmpWorkerIds(twi) & ")"
        Next twi
        brigNames(bDisp) = tmpWStr
    Next bDisp
    
    If brigCount = 1 Then
        selBrig(1) = True
    Else
        Dim dlg As Object
        Set dlg = ThisWorkbook.DialogSheets.Add
        dlg.DialogFrame.Width = 1000
        dlg.DialogFrame.Height = 500
        dlg.DialogFrame.Characters.Text = ""
        
        Dim lblDesc As Object
        Set lblDesc = dlg.Labels.Add(dlg.DialogFrame.Left + 15, dlg.DialogFrame.Top + 15, dlg.DialogFrame.Width - 120, 35)
        lblDesc.Text = UW(1042, 1099, 1073, 1077, 1088, 1080, 1090, 1077, 32, 1073, 1088, 1080, 1075, 1072, 1076, 1099, 32, 1085, 1072, 32) & Format(CDate(CDbl(curDateKey)), "dd"".""mm"".""yyyy") & UW(58, 10, 40, 1050, 1083, 1080, 1082, 1085, 1080, 1090, 1077, 32, 1087, 1086, 32, 1089, 1090, 1088, 1086, 1082, 1072, 1084, 32, 1076, 1083, 1103, 32, 1074, 1099, 1073, 1086, 1088, 1072, 32, 1085, 1077, 1089, 1082, 1086, 1083, 1100, 1082, 1080, 1093, 41)
        
        Dim dBtn As Object
        Dim btnTop As Double
        btnTop = dlg.DialogFrame.Top + 20
        For Each dBtn In dlg.Buttons
            dBtn.Left = dlg.DialogFrame.Left + dlg.DialogFrame.Width - dBtn.Width - 15
            dBtn.Top = btnTop
            btnTop = btnTop + dBtn.Height + 10
        Next dBtn
        
        Dim lb As Object
        Set lb = dlg.ListBoxes.Add(dlg.DialogFrame.Left + 15, dlg.DialogFrame.Top + 55, dlg.DialogFrame.Width - 120, dlg.DialogFrame.Height - 100)
        lb.MultiSelect = xlSimple
        
        Dim btnSelAll As Object
        Set btnSelAll = dlg.Buttons.Add(dlg.DialogFrame.Left + 15, dlg.DialogFrame.Top + dlg.DialogFrame.Height - 35, 120, 22)
        btnSelAll.Caption = UW(1042, 1099, 1073, 1088, 1072, 1090, 1100, 32, 1074, 1089, 1077)
        btnSelAll.OnAction = "SelectAllBrigades"
        
        Dim btnClearAll As Object
        Set btnClearAll = dlg.Buttons.Add(dlg.DialogFrame.Left + 145, dlg.DialogFrame.Top + dlg.DialogFrame.Height - 35, 120, 22)
        btnClearAll.Caption = UW(1054, 1095, 1080, 1089, 1090, 1080, 1090, 1100, 32, 1074, 1099, 1073, 1086, 1088)
        btnClearAll.OnAction = "ClearAllBrigades"
        
        For bDisp = 1 To brigCount
            lb.AddItem brigNames(bDisp)
        Next bDisp
        
        Application.ScreenUpdating = True
        Dim dlgResult As Boolean
        dlgResult = dlg.Show
        Application.ScreenUpdating = False
        
        Dim anySel As Boolean: anySel = False
        If dlgResult Then
            For bDisp = 1 To brigCount
                If lb.Selected(bDisp) Then
                    selBrig(bDisp) = True
                    anySel = True
                End If
            Next bDisp
        Else
            Application.DisplayAlerts = False
            dlg.Delete
            Application.DisplayAlerts = True
            wsMRS.Rows(outRow).Clear
            wsMRS.Rows(outRow).RowHeight = wsMRS.StandardHeight
            outRow = outRow - 1
            GoTo Done
        End If
        
        If Not anySel Then
            Application.DisplayAlerts = False
            dlg.Delete
            Application.DisplayAlerts = True
            wsMRS.Rows(outRow).Clear
            wsMRS.Rows(outRow).RowHeight = wsMRS.StandardHeight
            outRow = outRow - 1
            GoTo SkipDate
        End If
        
        Application.DisplayAlerts = False
        dlg.Delete
        Application.DisplayAlerts = True
    End If

    stage = "Write brigade blocks for date " & CStr(d)
    Dim b As Long
    For b = 1 To brigCount
        If Not selBrig(b) Then GoTo SkipBrigade

        Dim curBrigKey As String
        curBrigKey = brigKeys(b)

        Dim wStr As String
        wStr = brigNames(b)

        stage = "Write brigade header " & CStr(d) & "." & CStr(b)
        outRow = outRow + 1
        wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Merge
        wsMRS.Cells(outRow, 2).Value = UW(1041, 1088, 1080, 1075, 1072, 1076, 1072) & ": " & wStr
        SetCellBoldSafe wsMRS.Cells(outRow, 2), True, "MRS brigade header"
        wsMRS.Cells(outRow, 2).Font.Size = 16
        wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
        wsMRS.Cells(outRow, 2).HorizontalAlignment = -4108
        wsMRS.Cells(outRow, 2).Interior.Color = clrMrsSub
        wsMRS.Cells(outRow, 2).WrapText = True
        wsMRS.Rows(outRow).RowHeight = 60

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
        If ordInBrig > 1 Then QuickSortDoubleString ordTimes, ordKeys, 1, ordInBrig

        Dim dictWorkerLastRow As Object
        Set dictWorkerLastRow = CreateObject("Scripting.Dictionary")

        Dim oi As Long
        Dim pauseCellRow As Long

        For oi = 1 To ordInBrig
            Dim curOrder As String
            curOrder = ordKeys(oi)

            Dim blkIdx() As Long, blkCnt As Long
            ReDim blkIdx(1 To totalRows)
            blkCnt = 0
            For i = 1 To totalRows
                If arrOrder(i) = curOrder And CLng(arrSDate(i)) = curDateKey Then
                    blkCnt = blkCnt + 1
                    blkIdx(blkCnt) = i
                End If
            Next i

            If blkCnt > 1 Then QuickSortIndices blkIdx, arrSDate, arrSTime, arrOp, arrWID, 1, blkCnt

            Dim fIdx As Long
            fIdx = blkIdx(1)

            If oi > 1 Then
                stage = "Write pause row " & CStr(d) & "." & CStr(b) & "." & CStr(oi)
                outRow = outRow + 1
                pauseCellRow = outRow
                wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 8)).Merge
                wsMRS.Cells(outRow, 2).Value = UW(1055, 1072, 1091, 1079, 1072, 32, 40, 1084, 1080, 1085, 41, 58)
                SetCellBoldSafe wsMRS.Cells(outRow, 2), True, "MRS pause row"
                wsMRS.Cells(outRow, 2).HorizontalAlignment = -4108
                wsMRS.Cells(outRow, 9).NumberFormat = "0.00"
                wsMRS.Cells(outRow, 9).Value = 0
                wsMRS.Cells(outRow, 9).Locked = False
                If clrEditHasColor Then
                    wsMRS.Cells(outRow, 9).Interior.Color = clrEditable
                Else
                    wsMRS.Cells(outRow, 9).Interior.Pattern = xlNone
                End If
                ApplyDecimalValidation wsMRS.Cells(outRow, 9), 0, 999999.99
                wsMRS.Range(wsMRS.Cells(outRow, 10), wsMRS.Cells(outRow, 14)).Merge
                wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
            End If

            stage = "Write order header " & CStr(d) & "." & CStr(b) & "." & CStr(oi)
            outRow = outRow + 1
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 4)).Merge
            wsMRS.Cells(outRow, 2).Value = UW(1053, 1045, 32, 1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1086)
            With wsMRS.Cells(outRow, 2).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=UW(1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1086) & "," & UW(1053, 1045, 32, 1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1086)
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 4)).Locked = False

            wsMRS.Cells(outRow, 5).Value = UW(1047, 1072, 1082, 1072, 1079) & ":"
            wsMRS.Range(wsMRS.Cells(outRow, 6), wsMRS.Cells(outRow, 7)).Merge
            wsMRS.Range(wsMRS.Cells(outRow, 6), wsMRS.Cells(outRow, 7)).NumberFormat = "0"
            wsMRS.Cells(outRow, 6).Value = curOrder
            wsMRS.Range(wsMRS.Cells(outRow, 8), wsMRS.Cells(outRow, 14)).Merge
            wsMRS.Cells(outRow, 8).Value = UW(1056, 1072, 1073, 1086, 1095, 1077, 1077, 32, 1084, 1077, 1089, 1090, 1086) & ": " & arrWPlace(fIdx)

            wsMRS.Range(wsMRS.Cells(outRow, 5), wsMRS.Cells(outRow, 14)).Font.Size = 15
            SetCellBoldSafe wsMRS.Cells(outRow, 5), True, "MRS order label"
            SetCellBoldSafe wsMRS.Cells(outRow, 6), True, "MRS order number"
            SetCellBoldSafe wsMRS.Cells(outRow, 8), True, "MRS workplace"
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).HorizontalAlignment = -4108

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

            outRow = outRow + 1
            wsMRS.Cells(outRow, 10).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 40, 1095, 1072, 1089, 41)
            wsMRS.Cells(outRow, 11).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 40, 1084, 1080, 1085, 41)
            wsMRS.Cells(outRow, 12).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 40, 1095, 1072, 1089, 41)
            wsMRS.Cells(outRow, 13).Value = UW(1056, 1072, 1073, 1086, 1090, 1072, 40, 1084, 1080, 1085, 41)

            SetRangeBoldSafe wsMRS.Range(wsMRS.Cells(hdrRow1, 2), wsMRS.Cells(outRow, 14)), True, "MRS column header block"
            wsMRS.Range(wsMRS.Cells(hdrRow1, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
            wsMRS.Range(wsMRS.Cells(hdrRow1, 2), wsMRS.Cells(outRow, 14)).WrapText = True

            Dim dictWorkerSeenInOrder As Object
            Set dictWorkerSeenInOrder = CreateObject("Scripting.Dictionary")

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

                Dim rawDT As String
                If dictWorkerLastRow.Exists(curWID) Then
                    Dim prevR As Long
                    prevR = CLng(dictWorkerLastRow(curWID))
                    If Not dictWorkerSeenInOrder.Exists(curWID) Then
                        rawDT = "(H" & prevR & "+I" & prevR & "+I" & pauseCellRow & "/1440)"
                    Else
                        rawDT = "(H" & prevR & "+I" & prevR & ")"
                    End If
                    Dim shiftF As String
                    shiftF = BuildLunchShiftFormulaStr(rawDT, lunch1, lunch2, lunchDurMin)
                    wsMRS.Cells(outRow, 6).Formula = "=INT(" & shiftF & ")"
                    wsMRS.Cells(outRow, 7).Formula = "=MOD(" & shiftF & ",1)"
                Else
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
                    ApplyTimeValidation wsMRS.Cells(outRow, 7)
                End If
                wsMRS.Cells(outRow, 6).NumberFormat = "dd"".""mm"".""yyyy"
                wsMRS.Cells(outRow, 7).NumberFormat = "h:mm:ss"

                wsMRS.Cells(outRow, 10).Value = arrDur(idx)
                wsMRS.Cells(outRow, 11).Value = arrDur(idx) * 60

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
                wsMRS.Cells(outRow, 13).Locked = False
                If clrEditHasColor Then
                    wsMRS.Cells(outRow, 13).Interior.Color = clrEditable
                Else
                    wsMRS.Cells(outRow, 13).Interior.Pattern = xlNone
                End If
                ApplyDecimalValidation wsMRS.Cells(outRow, 13), 0, 999999.99

                Dim rStr As String: rStr = CStr(outRow)
                wsMRS.Cells(outRow, 12).Formula = "=M" & rStr & "/60"
                wsMRS.Cells(outRow, 12).NumberFormat = "0.00"

                Dim endD As String, endT As String
                BuildEndFormulasStr "F" & rStr, "G" & rStr, "M" & rStr, "1440", lunch1, lunch2, lunchDurMin, endD, endT
                wsMRS.Cells(outRow, 8).Formula = "=" & endD
                wsMRS.Cells(outRow, 8).NumberFormat = "dd"".""mm"".""yyyy"
                wsMRS.Cells(outRow, 9).Formula = "=" & endT
                wsMRS.Cells(outRow, 9).NumberFormat = "h:mm:ss"

                Dim iconF As String
                iconF = BuildLunchIconFormulaStr("G" & rStr, "M" & rStr, "1440", lunch1, lunch2, lunchDurMin, daStr)
                wsMRS.Cells(outRow, 14).Formula = "=" & iconF

                wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
                totalDur = totalDur + arrDur(idx)

                dictWorkerSeenInOrder(curWID) = True
                If dictWorkerLastRow.Exists(curWID) Then
                    dictWorkerLastRow(curWID) = outRow
                Else
                    dictWorkerLastRow.Add curWID, outRow
                End If
            Next i

            stage = "Write totals row " & CStr(d) & "." & CStr(b) & "." & CStr(oi)
            Dim dataEndRow As Long: dataEndRow = outRow
            outRow = outRow + 1
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 9)).Merge
            wsMRS.Cells(outRow, 2).Value = UW(1048, 1090, 1086, 1075, 1086) & ": " & blkCnt & " " & UW(1079, 1072, 1087, 1080, 1089, 1077, 1081)
            SetCellBoldSafe wsMRS.Cells(outRow, 2), True, "MRS totals label"
            wsMRS.Cells(outRow, 2).HorizontalAlignment = -4108
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

        outRow = outRow + 1
SkipBrigade:
    Next b

SkipDate:
    Next d

    Application.StatusBar = UW(1055, 1072, 1088, 1089, 1080, 1085, 1075, 32, 77, 82, 83, 58, 32, 1047, 1072, 1074, 1077, 1088, 1096, 1077, 1085, 1080, 1077, 32, 1092, 1086, 1088, 1084, 1072, 1090, 1080, 1088, 1086, 1074, 1072, 1085, 1080, 1103, 46, 46, 46)
    stage = "Finalize MRS formatting"
    Dim dataRange As Range
    Dim rr As Long
    Dim cellVal As String
    Dim mergeColCount As Long

    If outRow >= 4 Then
        Set dataRange = wsMRS.Range(wsMRS.Cells(4, 2), wsMRS.Cells(outRow, 14))
        dataRange.Font.Size = 14
        dataRange.Font.Color = GetContrastColor(clrLocked)
        dataRange.HorizontalAlignment = -4108
        dataRange.VerticalAlignment = -4108

        For rr = 4 To outRow
            If wsMRS.Cells(rr, 2).MergeCells Then
                mergeColCount = wsMRS.Cells(rr, 2).MergeArea.Columns.Count
                cellVal = CStr(wsMRS.Cells(rr, 2).Value)
                If mergeColCount >= 10 Then
                    If Left$(cellVal, 1) >= "0" And Left$(cellVal, 1) <= "9" Then
                        wsMRS.Cells(rr, 2).Font.Size = 18
                    ElseIf Left$(cellVal, 1) = ChrW(1041) Then
                        wsMRS.Cells(rr, 2).Font.Size = 16
                    End If
                End If
            End If
            cellVal = CStr(wsMRS.Cells(rr, 5).Value)
            If cellVal = UW(1047, 1072, 1082, 1072, 1079) & ":" Then
                wsMRS.Range(wsMRS.Cells(rr, 5), wsMRS.Cells(rr, 14)).Font.Size = 15
            End If
        Next rr
    End If

    RefreshMRSSheetColors wsMRS, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrMrsOrderUnconf, clrHeader

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
    ThisWorkbook.Protect UW(49, 49, 52, 55, 48, 57), True, False
    wsMRS.Protect UW(49, 49, 52, 55, 48, 57)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = calcState
    Application.StatusBar = False
    On Error GoTo 0
    If Len(errMsg) > 0 Then
        MsgBox UW(1054, 1096, 1080, 1073, 1082, 1072, 58, 32) & errMsg, vbCritical
    End If
    Exit Sub
EH:
    errMsg = stage & " | " & Err.Description
    Resume Done
End Sub

Private Sub QuickSortLong(arr() As Long, ByVal first As Long, ByVal last As Long)
    Dim pivot As Long, temp As Long
    Dim i As Long, j As Long
    If first >= last Then Exit Sub
    pivot = arr((first + last) \ 2)
    i = first: j = last
    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop
        If i <= j Then
            temp = arr(i): arr(i) = arr(j): arr(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then QuickSortLong arr, first, j
    If i < last Then QuickSortLong arr, i, last
End Sub

Private Sub QuickSortVariantStrings(arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim pivot As String, temp As Variant
    Dim i As Long, j As Long
    If first >= last Then Exit Sub
    pivot = CStr(arr((first + last) \ 2))
    i = first: j = last
    Do While i <= j
        Do While CStr(arr(i)) < pivot: i = i + 1: Loop
        Do While CStr(arr(j)) > pivot: j = j - 1: Loop
        If i <= j Then
            temp = arr(i): arr(i) = arr(j): arr(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then QuickSortVariantStrings arr, first, j
    If i < last Then QuickSortVariantStrings arr, i, last
End Sub

Private Sub QuickSortDoubleString(arrD() As Double, arrS() As String, ByVal first As Long, ByVal last As Long)
    Dim pivotD As Double, tempD As Double, tempS As String
    Dim i As Long, j As Long
    If first >= last Then Exit Sub
    pivotD = arrD((first + last) \ 2)
    i = first: j = last
    Do While i <= j
        Do While arrD(i) < pivotD: i = i + 1: Loop
        Do While arrD(j) > pivotD: j = j - 1: Loop
        If i <= j Then
            tempD = arrD(i): arrD(i) = arrD(j): arrD(j) = tempD
            tempS = arrS(i): arrS(i) = arrS(j): arrS(j) = tempS
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then QuickSortDoubleString arrD, arrS, first, j
    If i < last Then QuickSortDoubleString arrD, arrS, i, last
End Sub

Private Sub QuickSortIndices(idx() As Long, arrSDate() As Date, arrSTime() As Date, arrOp() As String, arrWID() As String, ByVal first As Long, ByVal last As Long)
    Dim pivotValA As Double, pivotValB As String, pivotValC As String
    Dim i As Long, j As Long, temp As Long
    If first >= last Then Exit Sub
    Dim midIdx As Long
    midIdx = idx((first + last) \ 2)
    pivotValA = CDbl(arrSDate(midIdx)) + CDbl(arrSTime(midIdx))
    pivotValB = arrOp(midIdx)
    pivotValC = arrWID(midIdx)
    i = first: j = last
    Do While i <= j
        Do While CompareIndices(idx(i), pivotValA, pivotValB, pivotValC, arrSDate, arrSTime, arrOp, arrWID) < 0: i = i + 1: Loop
        Do While CompareIndices(idx(j), pivotValA, pivotValB, pivotValC, arrSDate, arrSTime, arrOp, arrWID) > 0: j = j - 1: Loop
        If i <= j Then
            temp = idx(i): idx(i) = idx(j): idx(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then QuickSortIndices idx, arrSDate, arrSTime, arrOp, arrWID, first, j
    If i < last Then QuickSortIndices idx, arrSDate, arrSTime, arrOp, arrWID, i, last
End Sub

Private Function CompareIndices(idxA As Long, pA As Double, pB As String, pC As String, _
    arrSDate() As Date, arrSTime() As Date, arrOp() As String, arrWID() As String) As Integer
    Dim vA As Double
    vA = CDbl(arrSDate(idxA)) + CDbl(arrSTime(idxA))
    If vA < pA Then
        CompareIndices = -1: Exit Function
    ElseIf vA > pA Then
        CompareIndices = 1: Exit Function
    End If
    If arrOp(idxA) < pB Then
        CompareIndices = -1: Exit Function
    ElseIf arrOp(idxA) > pB Then
        CompareIndices = 1: Exit Function
    End If
    If arrWID(idxA) < pC Then
        CompareIndices = -1: Exit Function
    ElseIf arrWID(idxA) > pC Then
        CompareIndices = 1: Exit Function
    End If
    CompareIndices = 0
End Function

Public Sub ClearMRS()
    If Not ZQ() Then
        MsgBox UW(1054, 1096, 1080, 1073, 1082, 1072, 33), vbCritical
        Exit Sub
    End If
    On Error GoTo EH
    Dim wsIn As Worksheet, wsMRS As Worksheet
    Set wsIn = ThisWorkbook.Worksheets(2)
    Set wsMRS = ThisWorkbook.Worksheets(5)
    wsIn.Unprotect UW(49, 49, 52, 55, 48, 57)
    EnsureColorSettings wsIn
    Dim clrLocked As Long
    Dim clrEditHasColor As Boolean, clrEditable As Long
    Dim clrMrsHeader As Long, clrMrsSub As Long, clrMrsOrder As Long, clrMrsOrderUnconf As Long
    Dim clrHeader As Long
    ReadAllColors wsIn, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrMrsOrderUnconf, clrHeader
    wsIn.Protect UW(49, 49, 52, 55, 48, 57)
    wsMRS.Unprotect UW(49, 49, 52, 55, 48, 57)
    ClearMRSArea wsMRS, clrLocked
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

Private Sub ClearMRSArea(ByVal wsMRS As Worksheet, ByVal clrLocked As Long)
    Dim lastRow As Long
    lastRow = wsMRS.UsedRange.Row + wsMRS.UsedRange.Rows.Count - 1
    If lastRow >= 4 Then
        Dim prevEvents As Boolean
        prevEvents = Application.EnableEvents
        Application.EnableEvents = False
        wsMRS.Rows("4:" & CStr(lastRow)).Delete Shift:=xlUp
        Application.EnableEvents = prevEvents
    End If
End Sub

Public Sub SelectAllBrigades()
    On Error Resume Next
    Dim dlg As Object
    Set dlg = ActiveSheet
    If TypeName(dlg) = "DialogSheet" Then
        Dim lb As Object
        Set lb = dlg.ListBoxes(1)
        Dim i As Long
        For i = 1 To lb.ListCount
            lb.Selected(i) = True
        Next i
    End If
End Sub

Public Sub ClearAllBrigades()
    On Error Resume Next
    Dim dlg As Object
    Set dlg = ActiveSheet
    If TypeName(dlg) = "DialogSheet" Then
        Dim lb As Object
        Set lb = dlg.ListBoxes(1)
        Dim i As Long
        For i = 1 To lb.ListCount
            lb.Selected(i) = False
        Next i
    End If
End Sub

Public Sub SaveHistorySheet()
    ExportSheet ThisWorkbook.Worksheets(4), UW(1048, 1089, 1090, 1086, 1088, 1080, 1103, 95, 1056, 1072, 1089, 1095, 1077, 1090, 1086, 1074)
End Sub

Public Sub SaveMRSSheet()
    ExportSheet ThisWorkbook.Worksheets(5), UW(1055, 1072, 1088, 1089, 1080, 1085, 1075, 95, 77, 82, 83)
End Sub

Private Sub ExportSheet(ByVal ws As Worksheet, ByVal defaultNamePrefix As String)
    On Error GoTo EH
    Dim savePath As Variant
    Dim defaultName As String
    Dim newWb As Workbook
    Dim wsDisclaimerExport As Worksheet
    Dim wsDataExport As Worksheet
    Dim shp As Shape
    
    defaultName = defaultNamePrefix & "_" & Format(Now, "dd-mm-yyyy") & ".xlsx"
    
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultName, _
        FileFilter:=UW(1060, 1072, 1081, 1083, 1099) & " Excel (*.xlsx), *.xlsx", _
        Title:=UW(1057, 1086, 1093, 1088, 1072, 1085, 1080, 1090, 1100, 32, 1083, 1080, 1089, 1090) _
    )
    
    If savePath = False Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ThisWorkbook.Worksheets(1).Copy
    Set newWb = ActiveWorkbook
    Set wsDisclaimerExport = newWb.Worksheets(1)

    ws.Copy After:=newWb.Worksheets(newWb.Worksheets.Count)
    Set wsDataExport = newWb.Worksheets(newWb.Worksheets.Count)
    
    wsDisclaimerExport.Unprotect UW(49, 49, 52, 55, 48, 57)
    wsDataExport.Unprotect UW(49, 49, 52, 55, 48, 57)
    
    For Each shp In wsDataExport.Shapes
        shp.Delete
    Next shp

    Dim clrOrder As Long, clrUnconf As Long
    GetOrderColors clrOrder, clrUnconf

    Dim lastCol As Long
    If ws Is ThisWorkbook.Worksheets(4) Then lastCol = 22 Else lastCol = 14

    Dim lastDataRow As Long
    lastDataRow = LastContentRow(wsDataExport, 4)
    If lastDataRow >= 4 Then
        Dim cfRange As Range
        Set cfRange = wsDataExport.Range(wsDataExport.Cells(4, 2), wsDataExport.Cells(lastDataRow, lastCol))

        Dim sConf As String, sUnconf As String
        sConf = UW(1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1086)
        sUnconf = UW(1053, 1045, 32, 1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1086)

        With cfRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=$B4=""" & sConf & """")
            .Interior.Color = clrOrder
            .Font.Color = GetContrastColor(clrOrder)
            .StopIfTrue = True
        End With

        With cfRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=$B4=""" & sUnconf & """")
            .Interior.Color = clrUnconf
            .Font.Color = GetContrastColor(clrUnconf)
            .StopIfTrue = True
        End With
    End If

    wsDisclaimerExport.Protect UW(49, 49, 52, 55, 48, 57), True, True, False, False
    wsDataExport.Protect UW(49, 49, 52, 55, 48, 57), True, True, False, False
    wsDisclaimerExport.Activate
    
    newWb.SaveAs Filename:=CStr(savePath), FileFormat:=51
    newWb.Close SaveChanges:=False
    
    MsgBox UW(1060, 1072, 1081, 1083, 32, 1091, 1089, 1087, 1077, 1096, 1085, 1086, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 33), vbInformation
    
Cleanup:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
EH:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
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

    Dim syncOpsFromSelectors As Boolean
    syncOpsFromSelectors = False

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
    Set kmRange = Intersect(Target, Me.Range("K4:K23,N4:N23"))
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
        syncOpsFromSelectors = True
    End If

    If Not Intersect(Target, Me.Range("M4")) Is Nothing Then
        syncOpsFromSelectors = True
    End If

    If Not Intersect(Target, Me.Range("O4")) Is Nothing Then
        syncOpsFromSelectors = True
    End If

    If syncOpsFromSelectors Then
        Dim currentOpCount As Long
        currentOpCount = CLng(Val(Me.Range("B8").Value))
        If currentOpCount < 1 Then currentOpCount = 1
        If currentOpCount > 20 Then currentOpCount = 20
        SyncOperationRows Me, currentOpCount
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

    Application.EnableEvents = False
    Me.Unprotect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)

    Dim bRange As Range
    Set bRange = Intersect(Target, Me.Columns(2))
    If Not bRange Is Nothing Then
        Dim clrMrsOrder As Long, clrMrsOrderUnconf As Long
        GetOrderColors clrMrsOrder, clrMrsOrderUnconf
        Dim bCell As Range
        For Each bCell In bRange.Cells
            If bCell.Row >= 4 Then
                ColorOrderRow Me, bCell.Row, 22, clrMrsOrder, clrMrsOrderUnconf
            End If
        Next bCell
    End If

    Dim lRange As Range
    Set lRange = Intersect(Target, Union(Me.Columns(12), Me.Columns(5)))
    If Not lRange Is Nothing Then
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
            cell.NumberFormat = "0.00"
NextLCell:
        Next cell
    End If

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

    Application.EnableEvents = False
    Me.Unprotect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)

    Dim bRange As Range
    Set bRange = Intersect(Target, Me.Columns(2))
    If Not bRange Is Nothing Then
        Dim clrMrsOrder As Long, clrMrsOrderUnconf As Long
        GetOrderColors clrMrsOrder, clrMrsOrderUnconf
        Dim bCell As Range
        For Each bCell In bRange.Cells
            If bCell.Row >= 4 Then
                ColorOrderRow Me, bCell.Row, 14, clrMrsOrder, clrMrsOrderUnconf
            End If
        Next bCell
    End If

    If Target.Cells.Count = 1 And Target.Row >= 4 Then
        Dim col As Long
        col = Target.Column
        If (col = 13 Or col = 9) And Not Target.HasFormula Then
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
        End If
    End If

SafeExit:
    Me.Protect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)
    Application.EnableEvents = True
End Sub
'@

$workbookCode = @'
Private Sub Workbook_Open()
    Dim wsDisclaimer As Worksheet
    Dim wsIn As Worksheet
    Set wsDisclaimer = ThisWorkbook.Worksheets(1)
    Set wsIn = ThisWorkbook.Worksheets(2)
    wsIn.Unprotect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)
    wsIn.Range("B5").Value = Date
    wsIn.Range("B7").Value = Date
    wsIn.Protect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)
    RefreshWorkbookColors
    wsDisclaimer.Activate
    wsDisclaimer.Range("B2").Select
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

    while ($wb.Worksheets.Count -lt 5) {
        $null = $wb.Worksheets.Add()
    }

    $wsDisclaimer = $wb.Worksheets.Item(1)
    $wsDisclaimer.Name = (RU @(1044,1080,1089,1082,1083,1077,1081,1084,1077,1088))
    $wsInput = $wb.Worksheets.Item(2)
    $wsInput.Name = (RU @(1042,1074,1086,1076))
    $wsResult = $wb.Worksheets.Item(3)
    $wsResult.Name = (RU @(1056,1077,1079,1091,1083,1100,1090,1072,1090))
    $wsHistory = $wb.Worksheets.Item(4)
    $wsHistory.Name = (RU @(1048,1089,1090,1086,1088,1080,1103))
    $wsTechCards = $wb.Worksheets.Item(5)
    $wsTechCards.Name = (RU @(1055,1072,1088,1089,1080,1085,1075,32,77,82,83))

    # --- Title row ---
    $wsInput.Range("A1:P1").Merge() | Out-Null
    $wsInput.Range("A1").Value = (RU @(84,105,109,101,84,111,84,97,98,108,101,45,86,66,65,32,98,121,32,1043,1072,1083,1080,1084,1079,1103,1085,1086,1074,32,1043,46,1056,46))
    $wsInput.Range("A1").Font.Bold = $true
    $wsInput.Range("A1").Font.Size = 16
    $wsInput.Range("A1:P1").HorizontalAlignment = -4108  # xlCenter
    $wsInput.Range("A1:P1").VerticalAlignment = -4108     # xlCenter
    $wsInput.Rows(1).RowHeight = 50
    # --- Separator row: merged A2:P2 ---
    $wsInput.Range("A2:P2").Merge() | Out-Null

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
            $wsInput.Cells.Item($row, 4).Font.Color = $lightRed
            $wsInput.Cells.Item($row, 5).Font.Color = $lightRed
            $wsInput.Cells.Item($row, 5).Interior.Color = $lightRed
            $wsInput.Cells.Item($row, 5).Borders.LineStyle = -4142
        }
    }

    $wsInput.Range("G3").Value = (RU @(8470))
    $wsInput.Range("H3").Value = (RU @(1055,1044,1058,1042))
    $wsInput.Range("I3").Value = (RU @(1054,1087,1077,1088,1072,1094,1080,1103))
    $wsInput.Range("J3").Value = "-"
    $wsInput.Range("K3").Value = (RU @(1044,1083,1080,1090,1077,1083,1100,1085,1086,1089,1090,1100))
    $wsInput.Range("L3").Value = (RU @(1045,1076,1080,1085,1080,1094,1072))
    $wsInput.Range("M3").Value = (RU @(1058,1080,1087,32,1079,1072,1076,1072,1085,1080,1103))
    $wsInput.Range("N3").Value = (RU @(1055,1072,1091,1079,1072))
    $wsInput.Range("O3").Value = (RU @(1045,1076,1080,1085,1080,1094,1072))
    $wsInput.Range("P3").Value = (RU @(1059,1095,1072,1089,1090,1085,1080,1082,1080))
    $wsInput.Range("G3:P3").Font.Bold = $true

    $wsInput.Range("G4").Value2 = 1
    $wsInput.Range("I4").Value = (RU @(1054,1087,1077,1088,1072,1094,1080,1103,32,49))
    $wsInput.Range("K4").Value2 = 0
    $wsInput.Range("L4").Value = (RU @(1084,1080,1085))
    $wsInput.Range("M4").Value = (RU @(1054,1073,1097,1077,1077))
    $wsInput.Range("O4").Value = (RU @(1084,1080,1085))

    $wsInput.Range("G5").Value2 = 2
    $wsInput.Range("I5").Value = (RU @(1054,1087,1077,1088,1072,1094,1080,1103,32,50))
    $wsInput.Range("K5").Value2 = 0
    $wsInput.Range("L5").Value = (RU @(1084,1080,1085))
    $wsInput.Range("M5").Value = (RU @(1054,1073,1097,1077,1077))
    $wsInput.Range("O5").Value = (RU @(1084,1080,1085))

    # Default values: Пауза=0, Участники=пусто for active operations only
    for ($r = 4; $r -le 5; $r++) {
        $wsInput.Cells.Item($r, 14).Value2 = 0   # N = Пауза
        $wsInput.Cells.Item($r, 16).Value2 = ""  # P = Участники
    }

    # K4:K23, N4:N23: text format for safe input (NormalizeDecimal in VBA)
    $wsInput.Range("K4:K23").NumberFormat = "@"
    $wsInput.Range("N4:N23").NumberFormat = "@"

    $wsInput.Columns("A:A").ColumnWidth = 25
    $wsInput.Columns("B:B").ColumnWidth = 40
    $wsInput.Columns("C:C").ColumnWidth = 3
    $wsInput.Columns("D:D").ColumnWidth = 15
    $wsInput.Columns("E:E").ColumnWidth = 10
    $wsInput.Columns("F:F").ColumnWidth = 3
    $wsInput.Columns("G:G").ColumnWidth = 4
    $wsInput.Columns("H:H").ColumnWidth = 10
    $wsInput.Columns("I:I").ColumnWidth = 60
    $wsInput.Columns("J:J").ColumnWidth = 3
    $wsInput.Columns("K:K").ColumnWidth = 15
    $wsInput.Columns("L:L").ColumnWidth = 10
    $wsInput.Columns("M:M").ColumnWidth = 12
    $wsInput.Columns("N:N").ColumnWidth = 6
    $wsInput.Columns("O:O").ColumnWidth = 10
    $wsInput.Columns("P:P").ColumnWidth = 20
    $wsInput.Cells.HorizontalAlignment = -4108
    $wsInput.Cells.VerticalAlignment = -4108
    $wsInput.Cells.WrapText = $true

    # --- Light red for non-editable area ---
    $lightRed = 15132415  # RGB(255,230,230) in BGR
    $wsInput.Cells.Interior.Color = $lightRed

    # --- Disclaimer sheet ---
    $wsDisclaimer.Cells.Interior.Color = $lightRed
    $wsDisclaimer.Cells.Locked = $true
    $wsDisclaimer.Columns("B:B").ColumnWidth = 230
    $wsDisclaimer.Rows(2).RowHeight = 400
    $wsDisclaimer.Range("B2").Value = (RU @(1044,1040,1053,1053,1054,1045,32,1055,1056,1054,1043,1056,1040,1052,1052,1053,1054,1045,32,1054,1041,1045,1057,1055,1045,1063,1045,1053,1048,1045,40,1052,1040,1050,1056,1054,1057,41,32,1055,1056,1045,1044,1054,1057,1058,1040,1042,1051,1071,1045,1058,1057,1071,32,171,1050,1040,1050,32,1045,1057,1058,1068,187,44,32,1041,1045,1047,32,1050,1040,1050,1048,1061,45,1051,1048,1041,1054,32,1043,1040,1056,1040,1053,1058,1048,1049,44,32,1071,1042,1053,1054,32,1042,1067,1056,1040,1046,1045,1053,1053,1067,1061,32,1048,1051,1048,32,1055,1054,1044,1056,1040,1047,1059,1052,1045,1042,1040,1045,1052,1067,1061,44,32,1042,1050,1051,1070,1063,1040,1071,32,1043,1040,1056,1040,1053,1058,1048,1048,32,1058,1054,1042,1040,1056,1053,1054,1049,32,1055,1056,1048,1043,1054,1044,1053,1054,1057,1058,1048,44,32,1057,1054,1054,1058,1042,1045,1058,1057,1058,1042,1048,1071,32,1055,1054,32,1045,1043,1054,32,1050,1054,1053,1050,1056,1045,1058,1053,1054,1052,1059,32,1053,1040,1047,1053,1040,1063,1045,1053,1048,1070,32,1048,32,1054,1058,1057,1059,1058,1057,1058,1042,1048,1071,32,1053,1040,1056,1059,1064,1045,1053,1048,1049,44,32,1053,1054,32,1053,1045,32,1054,1043,1056,1040,1053,1048,1063,1048,1042,1040,1071,1057,1068,32,1048,1052,1048,46,32,1053,1048,32,1042,32,1050,1040,1050,1054,1052,32,1057,1051,1059,1063,1040,1045,32,1040,1042,1058,1054,1056,1067,32,1048,1051,1048,32,1055,1056,1040,1042,1054,1054,1041,1051,1040,1044,1040,1058,1045,1051,1048,32,1053,1045,32,1053,1045,1057,1059,1058,32,1054,1058,1042,1045,1058,1057,1058,1042,1045,1053,1053,1054,1057,1058,1048,32,1055,1054,32,1050,1040,1050,1048,1052,45,1051,1048,1041,1054,32,1048,1057,1050,1040,1052,44,32,1047,1040,32,1059,1065,1045,1056,1041,32,1048,1051,1048,32,1055,1054,32,1048,1053,1067,1052,32,1058,1056,1045,1041,1054,1042,1040,1053,1048,1071,1052,44,32,1042,32,1058,1054,1052,32,1063,1048,1057,1051,1045,44,32,1055,1056,1048,32,1044,1045,1049,1057,1058,1042,1048,1048,32,1050,1054,1053,1058,1056,1040,1050,1058,1040,44,32,1044,1045,1051,1048,1050,1058,1045,32,1048,1051,1048,32,1048,1053,1054,1049,32,1057,1048,1058,1059,1040,1062,1048,1048,44,32,1042,1054,1047,1053,1048,1050,1064,1048,1052,32,1048,1047,45,1047,1040,32,1048,1057,1055,1054,1051,1068,1047,1054,1042,1040,1053,1048,1071,32,1055,1056,1054,1043,1056,1040,1052,1052,1053,1054,1043,1054,32,1054,1041,1045,1057,1055,1045,1063,1045,1053,1048,1071,32,1048,1051,1048,32,1048,1053,1067,1061,32,1044,1045,1049,1057,1058,1042,1048,1049,32,1057,32,1055,1056,1054,1043,1056,1040,1052,1052,1053,1067,1052,32,1054,1041,1045,1057,1055,1045,1063,1045,1053,1048,1045,1052,46))
    $wsDisclaimer.Range("B2").Font.Bold = $true
    $wsDisclaimer.Range("B2").Font.Size = 28
    $wsDisclaimer.Range("B2").HorizontalAlignment = -4108  # xlCenter
    $wsDisclaimer.Range("B2").VerticalAlignment = -4108     # xlCenter
    $wsDisclaimer.Range("B2").WrapText = $true
    $wsDisclaimer.Protect((RU 49,49,52,55,48,57), $true, $true, $false, $false)

    # --- Clear color on always-editable cells ---
    $alwaysEditable = @(
        "B3:B17",
        "L4","M4","O4"
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
    # Editable op columns: H(8), I(9), J(10), K(11), N(14), P(16)
    $editOpCols = @(8, 9, 10, 11, 14, 16)
    # Synced locked columns: L(12), M(13), O(15)
    $syncOpCols = @(12, 13, 15)
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
                    # L4/M4/O4 are editable source cells
                    $wsInput.Cells.Item($row, $c).Font.Color = 0
                    $wsInput.Cells.Item($row, $c).Interior.Pattern = -4142
                    $wsInput.Cells.Item($row, $c).Locked = $false
                } else {
                    $wsInput.Cells.Item($row, $c).Interior.Color = $lightRed
                    $wsInput.Cells.Item($row, $c).Locked = $true
                    $wsInput.Cells.Item($row, $c).Font.Color = 0
                }
            }
        } else {
            for ($c = 7; $c -le 16; $c++) {
                $wsInput.Cells.Item($row, $c).Interior.Color = $lightRed
                $wsInput.Cells.Item($row, $c).Font.Color = $lightRed
                $wsInput.Cells.Item($row, $c).Borders.LineStyle = -4142
                $wsInput.Cells.Item($row, $c).Locked = $true
            }
        }
    }

    # Lock N4 (pause for first op) — no history in fresh file
    $wsInput.Range("N4").Locked = $true
    $wsInput.Range("N4").Interior.Color = $lightRed

    # --- L5:L23, M5:M23, O5:O23: synced values ---
    $minText = RU @(1084,1080,1085)
    $typeText = RU @(1054,1073,1097,1077,1077)
    $wsInput.Range("L5:L23").Value = $minText
    $wsInput.Range("M5:M23").Value = $typeText
    $wsInput.Range("O5:O23").Value = $minText

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

    # B5, B7: date
    foreach ($cell in @("B5","B7")) {
        $wsInput.Range($cell).Validation.Delete()
        $wsInput.Range($cell).Validation.Add(4, 1, 1, [int]([datetime]'2000-01-01').ToOADate(), [int]([datetime]'2099-12-31').ToOADate()) | Out-Null
        $wsInput.Range($cell).Validation.IgnoreBlank = $true
        $wsInput.Range($cell).Validation.ErrorTitle = $errTitle
    }

    # B6: time
    $wsInput.Range("B6").Validation.Delete()
    $wsInput.Range("B6").Validation.Add(5, 1, 1, 0, (86399 / 86400)) | Out-Null
    $wsInput.Range("B6").Validation.IgnoreBlank = $true
    $wsInput.Range("B6").Validation.ErrorTitle = $errTitle

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
    # K4:K23, N4:N23: decimal, 0-999999.99 (Data Validation + VBA normalization)
    foreach ($rng in @("K4:K23","N4:N23")) {
        $wsInput.Range($rng).Validation.Delete()
        $wsInput.Range($rng).Validation.Add(2, 1, 1, "0", $maxDec) | Out-Null
        $wsInput.Range($rng).Validation.IgnoreBlank = $true
        $wsInput.Range($rng).Validation.ErrorTitle = $errTitle
        $wsInput.Range($rng).Validation.ErrorMessage = $decErrMsg
    }

    # L4, O4: selector мин/час (only first row editable)
    $listSep = $excel.International(5)
    $unitList = (RU @(1084,1080,1085)) + $listSep + (RU @(1095,1072,1089))
    $unitErrMsg = (RU @(1042,1099,1073,1077,1088,1080,1090,1077,32,1084,1080,1085,32,1080,1083,1080,32,1095,1072,1089))  # "Выберите мин или час"
    foreach ($rng in @("L4","O4")) {
        $wsInput.Range($rng).Validation.Delete()
        $wsInput.Range($rng).Validation.Add(3, 1, 1, $unitList) | Out-Null
        $wsInput.Range($rng).Validation.IgnoreBlank = $true
        $wsInput.Range($rng).Validation.InCellDropdown = $true
        $wsInput.Range($rng).Validation.ErrorTitle = $errTitle
        $wsInput.Range($rng).Validation.ErrorMessage = $unitErrMsg
    }

    $typeList = (RU @(1054,1073,1097,1077,1077)) + $listSep + (RU @(1053,1072,32,1050,1072,1078,1076,1086,1075,1086))
    $typeErrMsg = (RU @(1042,1099,1073,1077,1088,1080,1090,1077,32,1090,1080,1087,32,1079,1072,1076,1072,1085,1080,1103))
    $wsInput.Range("M4").Validation.Delete()
    $wsInput.Range("M4").Validation.Add(3, 1, 1, $typeList) | Out-Null
    $wsInput.Range("M4").Validation.IgnoreBlank = $true
    $wsInput.Range("M4").Validation.InCellDropdown = $true
    $wsInput.Range("M4").Validation.ErrorTitle = $errTitle
    $wsInput.Range("M4").Validation.ErrorMessage = $typeErrMsg

    # P4:P23: text, max 100 chars
    $wsInput.Range("P4:P23").Validation.Delete()
    $wsInput.Range("P4:P23").Validation.Add(6, 1, 8, "100") | Out-Null
    $wsInput.Range("P4:P23").Validation.IgnoreBlank = $true
    $wsInput.Range("P4:P23").Validation.ErrorTitle = $errTitle
    $wsInput.Range("P4:P23").Validation.ErrorMessage = (RU @(1052,1072,1082,1089,1080,1084,1091,1084,32,49,48,48,32,1089,1080,1084,1074,1086,1083,1086,1074))

    $wsResult.Range("A1").Value2 = ""
    $wsResult.Range("B1").Value = (RU @(8470))
    $wsResult.Range("C1").Value = (RU @(1054,1087,1077,1088,1072,1094,1080,1103))
    $wsResult.Range("D1").Value = (RU @(1054,1073,1077,1076,63))
    $wsResult.Range("E1").Value = (RU @(1055,1072,1091,1079,1072,32,40,1084,1080,1085,41))
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

    $txt1 = RU @(1047,1072,1082,1072,1079,1086,1074,32,1085,1072,32,1083,1080,1089,1090,1077,58,32)
    $txt2 = RU @(32,124,32,1055,1086,1076,1090,1074,1077,1088,1078,1076,1077,1085,1086,58,32)
    $txt3 = RU @(32,124,32,1053,1045,32,1055,1086,1076,1090,1074,1077,1088,1078,1076,1077,1085,1086,58,32)
    $valY = RU @(1055,1086,1076,1090,1074,1077,1088,1078,1076,1077,1085,1086)
    $valN = RU @(1053,1045,32,1055,1086,1076,1090,1074,1077,1088,1078,1076,1077,1085,1086)
    $fHist = '="' + $txt1 + '" & (COUNTIF(B4:B10000, "' + $valY + '") + COUNTIF(B4:B10000, "' + $valN + '")) & "' + $txt2 + '" & COUNTIF(B4:B10000, "' + $valY + '") & "' + $txt3 + '" & COUNTIF(B4:B10000, "' + $valN + '")'

    $wsHistory.Range("B1:V1").Merge() | Out-Null
    $wsHistory.Range("B1").Value = (RU @(1048,1089,1090,1086,1088,1080,1103))
    $wsHistory.Range("B1").Formula = $fHist
    $wsHistory.Range("B1").Font.Bold = $true
    $wsHistory.Range("B1").Font.Size = 14
    $wsHistory.Range("B1:V1").Borders.LineStyle = 1  # xlContinuous
    $wsHistory.Range("B1:V1").HorizontalAlignment = -4108
    $wsHistory.Rows(1).RowHeight = 30

    $wsHistory.Range("B2:V2").Merge() | Out-Null
    $wsHistory.Range("B2").Value = (RU @(1057,1092,1086,1088,1084,1080,1088,1086,1074,1072,1085,1086,32,1087,1088,1080,32,1087,1086,1084,1086,1097,1080,32,1087,1088,1086,1075,1088,1072,1084,1084,1099,32,84,105,109,101,84,111,84,97,98,108,101,45,86,66,65,32,98,121,32,1043,1072,1083,1080,1084,1079,1103,1085,1086,1074,32,1043,46,1056,46))
    $wsHistory.Range("B2").Font.Bold = $true
    $wsHistory.Range("B2").Font.Size = 14
    $wsHistory.Range("B2:V2").Borders.LineStyle = 1
    $wsHistory.Rows(2).RowHeight = 30
    $wsHistory.Cells.Font.Size = 14
    
    $txt1 = RU @(1047,1072,1082,1072,1079,1086,1074,32,1085,1072,32,1083,1080,1089,1090,1077,58,32)
    $txt2 = RU @(32,124,32,1055,1086,1076,1090,1074,1077,1088,1078,1076,1077,1085,1086,58,32)
    $txt3 = RU @(32,124,32,1053,1045,32,1055,1086,1076,1090,1074,1077,1088,1078,1076,1077,1085,1086,58,32)
    $valY = RU @(1055,1086,1076,1090,1074,1077,1088,1078,1076,1077,1085,1086)
    $valN = RU @(1053,1045,32,1055,1086,1076,1090,1074,1077,1088,1078,1076,1077,1085,1086)
    $fHist = '="' + $txt1 + '" & (COUNTIF(B4:B10000, "' + $valY + '") + COUNTIF(B4:B10000, "' + $valN + '")) & "' + $txt2 + '" & COUNTIF(B4:B10000, "' + $valY + '") & "' + $txt3 + '" & COUNTIF(B4:B10000, "' + $valN + '")'

    $wsHistory.Range("B3:V3").Merge() | Out-Null
    $wsHistory.Range("B3").Formula = $fHist
    $wsHistory.Range("B3").Value = (RU @(1048,1089,1090,1086,1088,1080,1103))
    $wsHistory.Range("B3").Font.Bold = $true
    $wsHistory.Range("B3").Font.Size = 14
    $wsHistory.Range("B3:V3").Borders.LineStyle = 1
    $wsHistory.Range("B3:V3").HorizontalAlignment = -4108
    $wsHistory.Rows(3).RowHeight = 25
    $wsHistory.Columns("A:A").ColumnWidth = 2.7
    $wsHistory.Columns("B:B").ColumnWidth = 3.5
    $wsHistory.Columns("C:C").ColumnWidth = 60
    $wsHistory.Columns("D:D").ColumnWidth = 8
    $wsHistory.Columns("E:E").ColumnWidth = 14
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

    # --- History Save button ---
    $btnHistSaveLeft = [double]$wsHistory.Range("T1").Left + 10
    $btnHistSaveTop = [double]$wsHistory.Range("T1").Top + 4
    $btnHistSaveWidth = 100
    $btnHistSave = $wsHistory.Shapes.AddShape(1, $btnHistSaveLeft, $btnHistSaveTop, $btnHistSaveWidth, 22)
    $btnHistSave.Name = "btnSaveHistory"
    $btnHistSave.TextFrame.Characters().Text = (RU @(1057,1086,1093,1088,1072,1085,1080,1090,1100)) # "Сохранить"
    $btnHistSave.TextFrame.HorizontalAlignment = -4108  # xlCenter
    $btnHistSave.TextFrame.VerticalAlignment = -4108     # xlCenter
    $btnHistSave.Fill.ForeColor.RGB = 12419407  # RGB(79,129,189) blue
    $btnHistSave.Line.ForeColor.RGB = 3355443
    $btnHistSave.TextFrame.Characters().Font.Color = 16777215  # white
    $btnHistSave.OnAction = "SaveHistorySheet"

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

    $wsTechCards.Range("B1:N1").Merge() | Out-Null
    $wsTechCards.Range("B1").Formula = $fHist
    $wsTechCards.Range("B1").Font.Bold = $true
    $wsTechCards.Range("B1:N1").Borders.LineStyle = 1
    $wsTechCards.Range("B1:N1").HorizontalAlignment = -4108
    $wsTechCards.Rows(1).RowHeight = 30

    $wsTechCards.Range("B2:N2").Merge() | Out-Null
    $wsTechCards.Range("B2").Value = (RU @(1057,1092,1086,1088,1084,1080,1088,1086,1074,1072,1085,1086,32,1087,1088,1080,32,1087,1086,1084,1086,1097,1080,32,1087,1088,1086,1075,1088,1072,1084,1084,1099,32,84,105,109,101,84,111,84,97,98,108,101,45,86,66,65,32,98,121,32,1043,1072,1083,1080,1084,1079,1103,1085,1086,1074,32,1043,46,1056,46))
    $wsTechCards.Range("B2").Font.Bold = $true
    $wsTechCards.Range("B2").Font.Size = 14
    $wsTechCards.Range("B2:N2").Borders.LineStyle = 1
    $wsTechCards.Rows(2).RowHeight = 30

    $wsTechCards.Range("B3:N3").Merge() | Out-Null
    $wsTechCards.Range("B3").Value = (RU @(1055,1072,1088,1089,1080,1085,1075,32,77,82,83))  # Парсинг MRS
    $wsTechCards.Range("B3").Font.Bold = $true
    $wsTechCards.Range("B3:N3").Borders.LineStyle = 1
    $wsTechCards.Rows(3).RowHeight = 25

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

    # --- MRS Save button ---
    $btnMrsSaveLeft = [double]$wsTechCards.Range("L1").Left + 10
    $btnMrsSaveTop = [double]$wsTechCards.Range("L1").Top + 4
    $btnMrsSaveWidth = 100
    $btnMrsSave = $wsTechCards.Shapes.AddShape(1, $btnMrsSaveLeft, $btnMrsSaveTop, $btnMrsSaveWidth, 22)
    $btnMrsSave.Name = "btnSaveMRS"
    $btnMrsSave.TextFrame.Characters().Text = (RU @(1057,1086,1093,1088,1072,1085,1080,1090,1100)) # "Сохранить"
    $btnMrsSave.TextFrame.HorizontalAlignment = -4108  # xlCenter
    $btnMrsSave.TextFrame.VerticalAlignment = -4108     # xlCenter
    $btnMrsSave.Fill.ForeColor.RGB = 12419407  # RGB(79,129,189) blue
    $btnMrsSave.Line.ForeColor.RGB = 3355443
    $btnMrsSave.TextFrame.Characters().Font.Color = 16777215  # white
    $btnMrsSave.OnAction = "SaveMRSSheet"

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
    $wsInput.Range("A29:B29").Merge() | Out-Null
    $wsInput.Range("A29").Value = (RU @(1053,1072,1089,1090,1088,1086,1081,1082,1080,32,1062,1074,1077,1090,1086,1074))  # "Настройки Цветов"
    $wsInput.Range("A29").Font.Bold = $true

    $wsInput.Cells.Item(31, 1).Value = (RU @(1047,1072,1073,1083,1086,1082,1080,1088,1086,1074,1072,1085,1085,1099,1077))                         # "Заблокированные"
    $wsInput.Cells.Item(32, 1).Value = (RU @(1056,1077,1076,1072,1082,1090,1080,1088,1091,1077,1084,1099,1077))                                   # "Редактируемые"
    $wsInput.Cells.Item(33, 1).Value = (RU @(1055,1072,1088,1089,1080,1085,1075,32,77,82,83,32,1044,1072,1090,1072))                              # "Парсинг MRS Дата"
    $wsInput.Cells.Item(34, 1).Value = (RU @(1055,1072,1088,1089,1080,1085,1075,32,77,82,83,32,1041,1088,1080,1075,1072,1076,1072))               # "Парсинг MRS Бригада"
    $wsInput.Cells.Item(35, 1).Value = (RU @(1047,1072,1082,1072,1079,32,1055,1044,1058,1042))                                                      # "Заказ ПДТВ"
    $wsInput.Cells.Item(36, 1).Value = (RU @(1047,1072,1082,1072,1079,32,1053,1045,32,1055,1044,1058,1042))                                         # "Заказ НЕ ПДТВ"
    $wsInput.Cells.Item(37, 1).Value = (RU @(1064,1072,1087,1082,1072))                                                                            # "Шапка"

    # Lock merged header row 29
    $wsInput.Range("A29:B29").Locked = $true
    # Lock label cells (A31:A37), unlock color sample cells (B31:B37)
    for ($r = 31; $r -le 37; $r++) {
        $wsInput.Cells.Item($r, 1).Locked = $true
        $wsInput.Cells.Item($r, 2).Locked = $false
    }

    # Set default fill colors and borders in B column
    $clrDefLocked         = 15132415   # RGB(255,230,230)
    $clrDefMrsHeader      = 15189684   # RGB(180,198,231)
    $clrDefMrsSub         = 16768200   # RGB(200,220,255)
    $clrDefMrsOrder       = 13167560   # RGB(200,235,200)
    $clrDefMrsOrderUnconf = 14277081   # RGB(217,217,217)
    $clrDefHeader         = 13167560   # RGB(200,235,200)

    $wsInput.Cells.Item(31, 2).Interior.Color = $clrDefLocked
    # Row 32 (Editable) - no fill by default (xlNone)
    $wsInput.Cells.Item(32, 2).Interior.Pattern = -4142  # xlNone
    $wsInput.Cells.Item(33, 2).Interior.Color = $clrDefMrsHeader
    $wsInput.Cells.Item(34, 2).Interior.Color = $clrDefMrsSub
    $wsInput.Cells.Item(35, 2).Interior.Color = $clrDefMrsOrder
    $wsInput.Cells.Item(36, 2).Interior.Color = $clrDefMrsOrderUnconf
    $wsInput.Cells.Item(37, 2).Interior.Color = $clrDefHeader

    # Add borders to B31:B37
    for ($r = 31; $r -le 37; $r++) {
        $wsInput.Cells.Item($r, 2).Borders.LineStyle = 1  # xlContinuous
    }

    # --- Small "pick color" buttons in column C for rows 31-37 ---
    $pickBtnNames = @(
        @{ Row = 31; Action = "PickColorLocked" },
        @{ Row = 32; Action = "PickColorEditable" },
        @{ Row = 33; Action = "PickColorMrsHeader" },
        @{ Row = 34; Action = "PickColorMrsSub" },
        @{ Row = 35; Action = "PickColorMrsOrder" },
        @{ Row = 36; Action = "PickColorMrsOrderUnconf" },
        @{ Row = 37; Action = "PickColorHeader" }
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
        $wsHistory.Activate()
        $excel.ActiveWindow.SplitRow = 1
        $excel.ActiveWindow.FreezePanes = $true
        $wsTechCards.Activate()
        $excel.ActiveWindow.SplitRow = 1
        $excel.ActiveWindow.FreezePanes = $true
        $wsInput.Activate()
    } catch {}

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
    $wb.Protect("114709", $true, $false)
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
