Attribute VB_Name = "modTimeToTable"
Option Explicit

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
                    startDT = ShiftStartOutOfLunch(intendedStart, lunch1, lunch2, lunchDurMin)
                ElseIf isFirstWorkerInOp Then
                    intendedStart = prevEndDT + breakDays
                    startDT = ShiftStartOutOfLunch(intendedStart, lunch1, lunch2, lunchDurMin)
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

                wsOut.Cells(outRow, 22).Value = CStr(opNum) & "_" & CStr(workerValue)

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

    AppendResultToHistory wsHist, wsOut, wsIn, outRow - 1, zRow + 5, workerCount, lunch1, lunch2, lunchDurMin, primaryIsMin

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
    wsIn.Cells(4, 13).Interior.Pattern = xlNone
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
    On Error Resume Next
    Dim wsR As Worksheet, wsH As Worksheet
    Set wsR = ThisWorkbook.Worksheets(2)
    Set wsH = ThisWorkbook.Worksheets(3)
    wsR.Unprotect UW(49, 49, 52, 55, 48, 57)
    ClearResultArea wsR
    wsR.Protect UW(49, 49, 52, 55, 48, 57)
    wsH.Unprotect UW(49, 49, 52, 55, 48, 57)
    ClearHistoryArea wsH
    wsH.Protect UW(49, 49, 52, 55, 48, 57)
    On Error GoTo 0
End Sub

Private Sub ClearResultArea(ByVal wsOut As Worksheet)
    On Error Resume Next
    wsOut.Range("B2:V20000").UnMerge
    On Error GoTo 0
    wsOut.Range("B2:V20000").ClearContents
    wsOut.Range("B2:V20000").Interior.Color = RGB(255, 230, 230)
    wsOut.Range("B2:V20000").Borders.LineStyle = xlNone
    wsOut.Range("B2:V20000").HorizontalAlignment = -4108
End Sub

Private Sub ClearHistoryArea(ByVal wsHist As Worksheet)
    Dim rng As Range
    Application.EnableEvents = False
    On Error Resume Next
    wsHist.Range("B2:V20000").UnMerge
    On Error GoTo 0
    Set rng = wsHist.Range("B2:V20000")
    rng.ClearContents
    rng.Interior.Color = RGB(255, 230, 230)
    rng.Borders.LineStyle = xlNone
    rng.HorizontalAlignment = -4108
    rng.VerticalAlignment = -4108
    rng.Font.Size = 14
    rng.Font.Color = RGB(0, 0, 0)
    rng.Font.Bold = False
    rng.NumberFormat = "General"
    rng.Locked = True
    rng.WrapText = False
    wsHist.Rows("2:20000").RowHeight = wsHist.StandardHeight
    Application.EnableEvents = True
End Sub

Public Sub SyncWorkerIdInputs(ByVal wsIn As Worksheet, ByVal workerCount As Long)
    Dim i As Long

    If workerCount < 1 Then workerCount = 1
    If workerCount > 10 Then workerCount = 10

    wsIn.Range("D3").Value = UW(1048, 1089, 1087, 1086, 1083, 1085, 1080, 1090, 1077, 1083, 1100)
    wsIn.Range("D3:E3").Font.Bold = True
    wsIn.Range("E4:E13").NumberFormat = "@"

    For i = 1 To 10
        Dim rowNum As Long
        rowNum = 3 + i

        wsIn.Cells(rowNum, 4).Value = UW(1048, 1089, 1087, 1086, 1083, 1085, 1080, 1090, 1077, 1083, 1100, 32) & i
        StyleWorkerInputRow wsIn, rowNum, (i <= workerCount)
    Next i
End Sub

Public Sub SyncOperationRows(ByVal wsIn As Worksheet, ByVal opCount As Long)
    Application.EnableEvents = False
    On Error GoTo Cleanup

    If opCount < 1 Then opCount = 1
    If opCount > 20 Then opCount = 20

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
        wsIn.Cells(r, 7).Font.Color = RGB(0, 0, 0)
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
            wsIn.Cells(r, editCols(c)).Interior.Pattern = xlNone
            wsIn.Cells(r, editCols(c)).Borders.LineStyle = xlContinuous
            wsIn.Cells(r, editCols(c)).Font.Color = RGB(0, 0, 0)
        Next c
        ' Lock pause (M) for first operation if no history exists
        If i = 1 Then
            Dim wsHist As Worksheet
            Set wsHist = ThisWorkbook.Worksheets(3)
            Dim histLastRow As Long
            histLastRow = wsHist.Cells(wsHist.Rows.Count, 2).End(xlUp).Row
            If histLastRow <= 1 Then
                wsIn.Cells(r, 13).Locked = True
                wsIn.Cells(r, 13).Interior.Color = RGB(255, 230, 230)
            End If
        End If

        ' Synced columns: visible but locked (except row 2 = source)
        For c = LBound(syncCols) To UBound(syncCols)
            wsIn.Cells(r, syncCols(c)).Borders.LineStyle = xlContinuous
            wsIn.Cells(r, syncCols(c)).Font.Color = RGB(0, 0, 0)
            If i = 1 Then
                wsIn.Cells(r, syncCols(c)).Interior.Pattern = xlNone
                wsIn.Cells(r, syncCols(c)).Locked = False
            Else
                wsIn.Cells(r, syncCols(c)).Interior.Color = RGB(255, 230, 230)
                wsIn.Cells(r, syncCols(c)).Locked = True
            End If
        Next c
    Next i

    ' Lock and color unused rows
    If opCount + 4 <= 23 Then
        Dim unusedRange As Range
        Set unusedRange = wsIn.Range(wsIn.Cells(opCount + 4, 7), wsIn.Cells(23, 15))
        unusedRange.ClearContents
        unusedRange.Interior.Color = RGB(255, 230, 230)
        unusedRange.Font.Color = RGB(255, 230, 230)
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

Private Sub StyleWorkerInputRow(ByVal wsIn As Worksheet, ByVal rowNum As Long, ByVal isVisible As Boolean)
    With wsIn.Cells(rowNum, 4)
        If isVisible Then
            .Font.Color = RGB(0, 0, 0)
        Else
            .Font.Color = RGB(255, 230, 230)
            .ClearContents
        End If
        .Locked = True
    End With

    With wsIn.Cells(rowNum, 5)
        .NumberFormat = "00000000"
        If isVisible Then
            .Font.Color = RGB(0, 0, 0)
            .Interior.Pattern = xlNone
            .Borders.LineStyle = xlContinuous
            .Locked = False
        Else
            .Font.Color = RGB(255, 230, 230)
            .Interior.Color = RGB(255, 230, 230)
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
    ByRef opStart As Date, _
    ByVal durationDays As Double, _
    ByVal lunch1 As Date, _
    ByVal lunch2 As Date, _
    ByVal hasLunch2 As Boolean, _
    ByVal lunchDurDays As Double, _
    ByRef crossedLunch As Boolean) As Date

    Dim opEnd As Date, changed As Boolean, guard As Long
    opEnd = opStart + durationDays
    guard = 0

    Do
        changed = False
        guard = guard + 1
        If guard > 2000 Then Exit Do

        If CheckLunchWindow(opStart, opEnd, lunch1, lunchDurDays, crossedLunch) Then
            changed = True
            GoTo ContinueLoop
        End If

        If hasLunch2 Then
            If CheckLunchWindow(opStart, opEnd, lunch2, lunchDurDays, crossedLunch) Then
                changed = True
                GoTo ContinueLoop
            End If
        End If

ContinueLoop:
    Loop While changed

    ComputeEndWithLunch = opEnd
End Function

Public Function ShiftStartOutOfLunch( _
    ByVal startDateTime As Date, _
    ByVal lunch1Raw As Variant, _
    ByVal lunch2Raw As Variant, _
    ByVal lunchDurMin As Double) As Date

    On Error GoTo EH

    Dim lunch1 As Date, lunch2 As Date, hasLunch2 As Boolean, lunchDurDays As Double
    ParseLunchParams lunch1Raw, lunch2Raw, lunchDurMin, lunch1, lunch2, hasLunch2, lunchDurDays

    Dim endProbe As Date, crossed As Boolean
    endProbe = startDateTime
    crossed = False

    If CheckLunchWindow(startDateTime, endProbe, lunch1, lunchDurDays, crossed) Then
        ' startDateTime уже сдвинут внутри окна
    End If
    If hasLunch2 Then
        crossed = False
        endProbe = startDateTime
        If CheckLunchWindow(startDateTime, endProbe, lunch2, lunchDurDays, crossed) Then
            ' startDateTime уже сдвинут
        End If
    End If

    ShiftStartOutOfLunch = startDateTime
    Exit Function
EH:
    ShiftStartOutOfLunch = startDateTime
End Function

Public Function ComputeEndWithLunchFormula( _
    ByVal startDateTime As Date, _
    ByVal durationDays As Double, _
    ByVal lunch1Raw As Variant, _
    ByVal lunch2Raw As Variant, _
    ByVal lunchDurMin As Double) As Date

    On Error GoTo EH

    Dim lunch1 As Date, lunch2 As Date, hasLunch2 As Boolean, lunchDurDays As Double
    ParseLunchParams lunch1Raw, lunch2Raw, lunchDurMin, lunch1, lunch2, hasLunch2, lunchDurDays

    Dim st As Date, en As Date, crossed As Boolean
    st = startDateTime
    crossed = False
    en = ComputeEndWithLunch(st, durationDays, lunch1, lunch2, hasLunch2, lunchDurDays, crossed)
    ComputeEndWithLunchFormula = en
    Exit Function
EH:
    ComputeEndWithLunchFormula = startDateTime + durationDays
End Function

Public Function DidCrossLunch( _
    ByVal startDateTime As Date, _
    ByVal endDateTime As Date, _
    ByVal lunch1Raw As Variant, _
    ByVal lunch2Raw As Variant, _
    ByVal lunchDurMin As Double) As Boolean

    On Error GoTo EH

    Dim lunch1 As Date, lunch2 As Date, hasLunch2 As Boolean, lunchDurDays As Double
    ParseLunchParams lunch1Raw, lunch2Raw, lunchDurMin, lunch1, lunch2, hasLunch2, lunchDurDays

    DidCrossLunch = IntervalIntersectsLunch(startDateTime, endDateTime, lunch1, lunchDurDays)
    If Not DidCrossLunch And hasLunch2 Then
        DidCrossLunch = IntervalIntersectsLunch(startDateTime, endDateTime, lunch2, lunchDurDays)
    End If
    Exit Function
EH:
    DidCrossLunch = False
End Function

Public Function LunchFlag( _
    ByVal startDateTime As Date, _
    ByVal endDateTime As Date, _
    ByVal lunch1Raw As Variant, _
    ByVal lunch2Raw As Variant, _
    ByVal lunchDurMin As Double) As String

    If DidCrossLunch(startDateTime, endDateTime, lunch1Raw, lunch2Raw, lunchDurMin) Then
        LunchFlag = UW(1044, 1040)
    Else
        LunchFlag = ""
    End If
End Function

Private Function IntervalIntersectsLunch( _
    ByVal startDateTime As Date, _
    ByVal endDateTime As Date, _
    ByVal lunchTime As Date, _
    ByVal lunchDurDays As Double) As Boolean

    Dim daySerial As Long
    For daySerial = CLng(Int(startDateTime)) - 1 To CLng(Int(endDateTime)) + 1
        Dim lunchStart As Date, lunchEnd As Date
        lunchStart = daySerial + TimePart(lunchTime)
        lunchEnd = lunchStart + lunchDurDays

        If (startDateTime < lunchStart And endDateTime > lunchStart) Or _
           (startDateTime >= lunchStart And startDateTime < lunchEnd) Then
            IntervalIntersectsLunch = True
            Exit Function
        End If
    Next daySerial

    IntervalIntersectsLunch = False
End Function

Private Function CheckLunchWindow( _
    ByRef opStart As Date, _
    ByRef opEnd As Date, _
    ByVal lunchTime As Date, _
    ByVal lunchDurDays As Double, _
    ByRef crossedLunch As Boolean) As Boolean

    Dim durationKeep As Double
    durationKeep = opEnd - opStart

    Dim daySerial As Long
    For daySerial = CLng(Int(opStart)) - 1 To CLng(Int(opEnd)) + 1
        Dim lunchStart As Date, lunchEnd As Date
        lunchStart = daySerial + TimePart(lunchTime)
        lunchEnd = lunchStart + lunchDurDays

        If opStart >= lunchStart And opStart < lunchEnd Then
            opStart = lunchEnd
            opEnd = opStart + durationKeep
            crossedLunch = True
            CheckLunchWindow = True
            Exit Function
        End If

        If opStart < lunchStart And opEnd > lunchStart Then
            Dim beforeLunch As Double, afterLunch As Double
            beforeLunch = lunchStart - opStart
            afterLunch = durationKeep - beforeLunch
            If afterLunch < 0 Then afterLunch = 0

            opStart = lunchEnd
            opEnd = opStart + afterLunch
            durationKeep = opEnd - opStart
            crossedLunch = True
            CheckLunchWindow = True
            Exit Function
        End If
    Next daySerial

    CheckLunchWindow = False
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
    ByVal primaryIsMin As Boolean)

    Dim hLunch1 As String, hLunch2 As String, hLunchDur As String
    hLunch1 = """" & Format$(lunch1, "hh:nn:ss") & """"
    hLunch2 = """" & Format$(lunch2, "hh:nn:ss") & """"
    hLunchDur = NumForFormula(lunchDurMin)

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
        wsHist.Cells(dataStartRow, 18).Formula = "=INT(ShiftStartOutOfLunch(T" & prevDataEndRow & "+U" & prevDataEndRow & "+E" & dataStartRow & "," & hLunch1 & "," & hLunch2 & "," & hLunchDur & "))"
        wsHist.Cells(dataStartRow, 19).Formula = "=MOD(ShiftStartOutOfLunch(T" & prevDataEndRow & "+U" & prevDataEndRow & "+E" & dataStartRow & "," & hLunch1 & "," & hLunch2 & "," & hLunchDur & "),1)"
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
            wsHist.Cells(rr, 18).Formula = "=IF(B" & rr & "=B" & (rr - 1) & ",R" & (rr - 1) & ",INT(ShiftStartOutOfLunch(T" & (rr - 1) & "+U" & (rr - 1) & "+E" & rr & "," & hLunch1 & "," & hLunch2 & "," & hLunchDur & ")))"
            wsHist.Cells(rr, 19).Formula = "=IF(B" & rr & "=B" & (rr - 1) & ",S" & (rr - 1) & ",MOD(ShiftStartOutOfLunch(T" & (rr - 1) & "+U" & (rr - 1) & "+E" & rr & "," & hLunch1 & "," & hLunch2 & "," & hLunchDur & "),1))"
        End If

        wsHist.Cells(rr, 20).Formula = "=INT(ComputeEndWithLunchFormula(R" & rr & "+S" & rr & ",L" & rr & "/" & lDivisor & "," & hLunch1 & "," & hLunch2 & "," & hLunchDur & "))"
        wsHist.Cells(rr, 21).Formula = "=MOD(ComputeEndWithLunchFormula(R" & rr & "+S" & rr & ",L" & rr & "/" & lDivisor & "," & hLunch1 & "," & hLunch2 & "," & hLunchDur & "),1)"
    Next rr

    For rr = dataStartRow To dataEndRow
        wsHist.Cells(rr, 4).Formula = "=LunchFlag(R" & rr & "+S" & rr & ",T" & rr & "+U" & rr & "," & hLunch1 & "," & hLunch2 & "," & hLunchDur & ")"
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
                wsHist.Cells(rr, editCol(ec)).Interior.Pattern = xlNone
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
        wsHist.Cells(dataStartRow, 18).Interior.Pattern = xlNone
        wsHist.Cells(dataStartRow, 19).Locked = False
        wsHist.Cells(dataStartRow, 19).Interior.Pattern = xlNone
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

    Dim wsMRS As Worksheet
    Set wsMRS = ThisWorkbook.Worksheets(4)

    Dim filePath As Variant
    If MsgBox(UW(1042, 1099, 1075, 1088, 1091, 1078, 1072, 1081, 1090, 1077, 32, 1076, 1072, 1085, 1085, 1099, 1077, 32, 1079, 1072, 32, 1054, 1044, 1048, 1053, 32, 1076, 1077, 1085, 1100, 44, 32, 1080, 1085, 1072, 1095, 1077, 32, 1073, 1091, 1076, 1077, 1090, 32, 1073, 1077, 1076, 1072, 46, 46, 46), vbOKCancel + vbExclamation) = vbCancel Then Exit Sub

    filePath = Application.GetOpenFilename(UW(1060, 1072, 1081, 1083, 1099) & " Excel (*.xlsx), *.xlsx")
    If filePath = False Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    wsMRS.Unprotect UW(49, 49, 52, 55, 48, 57)
    ClearMRSArea wsMRS

    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(CStr(filePath), ReadOnly:=True)
    Dim srcWs As Worksheet
    Set srcWs = srcWb.Sheets(1)

    Dim lastRow As Long
    lastRow = srcWs.Cells(srcWs.Rows.Count, 5).End(xlUp).Row
    If lastRow < 2 Then
        srcWb.Close False
        MsgBox UW(1060, 1072, 1081, 1083, 32, 1087, 1091, 1089, 1090), vbInformation
        GoTo Done
    End If

    Dim data As Variant
    data = srcWs.Range("A1:Z" & lastRow).Value
    srcWb.Close False

    Dim totalRows As Long
    totalRows = UBound(data, 1) - 1

    ' Parse rows into arrays
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
        outRow = outRow + 1
        wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Merge
        wsMRS.Cells(outRow, 2).Value = Format(CDate(CDbl(curDateKey)), "dd"".""mm"".""yyyy")
        wsMRS.Cells(outRow, 2).Font.Bold = True
        wsMRS.Cells(outRow, 2).Font.Size = 18
        wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
        wsMRS.Cells(outRow, 2).HorizontalAlignment = -4108
        wsMRS.Cells(outRow, 2).Interior.Color = RGB(180, 198, 231)
        wsMRS.Rows(outRow).RowHeight = 40

        ' === Step 1: Build worker set per order (filtered by date) ===
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
        outRow = outRow + 1
        wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Merge
        wsMRS.Cells(outRow, 2).Value = UW(1041, 1088, 1080, 1075, 1072, 1076, 1072) & ": " & wStr
        wsMRS.Cells(outRow, 2).Font.Bold = True
        wsMRS.Cells(outRow, 2).Font.Size = 16
        wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
        wsMRS.Cells(outRow, 2).HorizontalAlignment = -4131
        wsMRS.Cells(outRow, 2).Interior.Color = RGB(200, 220, 255)
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
                outRow = outRow + 1
                pauseCellRow = outRow
                wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 8)).Merge
                wsMRS.Cells(outRow, 2).Value = UW(1055, 1072, 1091, 1079, 1072, 32, 40, 1084, 1080, 1085, 41, 58)
                wsMRS.Cells(outRow, 2).Font.Bold = True
                wsMRS.Cells(outRow, 2).HorizontalAlignment = -4152
                wsMRS.Cells(outRow, 9).NumberFormat = "0.00"
                wsMRS.Cells(outRow, 9).Value = 0
                wsMRS.Cells(outRow, 9).NumberFormat = "@"
                wsMRS.Cells(outRow, 9).Locked = False
                wsMRS.Cells(outRow, 9).Interior.Color = RGB(255, 255, 255)
                wsMRS.Range(wsMRS.Cells(outRow, 10), wsMRS.Cells(outRow, 14)).Merge
                wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
            End If

            ' Order header (merged B:N)
            outRow = outRow + 1
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Merge
            wsMRS.Cells(outRow, 2).Value = UW(1047, 1072, 1082, 1072, 1079) & ": " & curOrder & " | " & UW(1056, 1072, 1073, 1086, 1095, 1077, 1077, 32, 1084, 1077, 1089, 1090, 1086) & ": " & arrWPlace(fIdx)
            wsMRS.Cells(outRow, 2).Font.Bold = True
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
            wsMRS.Cells(outRow, 2).HorizontalAlignment = -4131

            ' Column headers row 1
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
            wsMRS.Range(wsMRS.Cells(hdrRow1, 2), wsMRS.Cells(outRow, 14)).Font.Bold = True
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
                rawDT = ""
                If dictWorkerLastRow.Exists(curWID) Then
                    Dim prevR As Long
                    prevR = CLng(dictWorkerLastRow(curWID))
                    If Not dictWorkerSeenInOrder.Exists(curWID) Then
                        ' Cross-order: previous end + pause
                        rawDT = "H" & prevR & "+I" & prevR & "+I" & pauseCellRow & "/1440"
                    Else
                        ' Within order: start = previous end for same worker
                        rawDT = "H" & prevR & "+I" & prevR
                    End If
                    wsMRS.Cells(outRow, 6).Formula = "=INT(ShiftStartOutOfLunch(" & rawDT & ",""12:00:00"",""00:00:00"",45))"
                    wsMRS.Cells(outRow, 7).Formula = "=MOD(ShiftStartOutOfLunch(" & rawDT & ",""12:00:00"",""00:00:00"",45),1)"
                Else
                    ' First time seeing this worker: parsed values (G editable, shift if in lunch)
                    wsMRS.Cells(outRow, 6).Value = arrSDate(idx)
                    Dim startTimeVal As Date
                    startTimeVal = arrSTime(idx)
                    If CDbl(startTimeVal) >= CDbl(TimeSerial(12, 0, 0)) And CDbl(startTimeVal) < CDbl(TimeSerial(12, 45, 0)) Then
                        startTimeVal = TimeSerial(12, 45, 0)
                    End If
                    wsMRS.Cells(outRow, 7).Value = startTimeVal
                    wsMRS.Cells(outRow, 7).Locked = False
                    wsMRS.Cells(outRow, 7).Interior.Color = RGB(255, 255, 255)
                End If
                wsMRS.Cells(outRow, 6).NumberFormat = "dd"".""mm"".""yyyy"
                wsMRS.Cells(outRow, 7).NumberFormat = "h:mm:ss"

                ' Parsing duration (static from source)
                wsMRS.Cells(outRow, 10).Value = arrDur(idx)
                wsMRS.Cells(outRow, 11).Value = arrDur(idx) * 60

                ' M = Расчетное мин (INPUT, editable) — initial value from parsed times
                Dim initMin As Double
                initMin = (CDbl(arrEDate(idx)) + CDbl(arrETime(idx)) - CDbl(arrSDate(idx)) - CDbl(arrSTime(idx))) * 1440
                If CDbl(arrSTime(idx)) <= CDbl(TimeSerial(12, 0, 0)) And CDbl(arrETime(idx)) > CDbl(TimeSerial(12, 0, 0)) And arrSDate(idx) = arrEDate(idx) Then
                    initMin = initMin - 45
                End If
                If initMin < 0 Then initMin = 0
                wsMRS.Cells(outRow, 13).NumberFormat = "0.00"
                wsMRS.Cells(outRow, 13).Value = Round(initMin, 2)
                wsMRS.Cells(outRow, 13).NumberFormat = "@"
                wsMRS.Cells(outRow, 13).Locked = False
                wsMRS.Cells(outRow, 13).Interior.Color = RGB(255, 255, 255)

                ' L = Расчетное час (formula from M)
                Dim r As String: r = CStr(outRow)
                wsMRS.Cells(outRow, 12).Formula = "=M" & r & "/60"
                wsMRS.Cells(outRow, 12).NumberFormat = "0.00"

                ' H = End date (formula: start + duration + lunch)
                ' I = End time (formula: start + duration + lunch)
                ' lunch_adj = IF(AND(G<12:00, G+M/1440>12:00), 45/1440, 0)
                Dim lunchF As String
                lunchF = "IF(AND(G" & r & "<=TIME(12,0,0),G" & r & "+M" & r & "/1440>TIME(12,0,0)),45/1440,0)"
                wsMRS.Cells(outRow, 8).Formula = "=INT(F" & r & "+G" & r & "+M" & r & "/1440+" & lunchF & ")"
                wsMRS.Cells(outRow, 8).NumberFormat = "dd"".""mm"".""yyyy"
                wsMRS.Cells(outRow, 9).Formula = "=MOD(G" & r & "+M" & r & "/1440+" & lunchF & ",1)"
                wsMRS.Cells(outRow, 9).NumberFormat = "h:mm:ss"

                ' N = Обед? (formula)
                If rawDT <> "" Then
                    ' Formula-based G: use raw (unshifted) datetime with LunchFlag UDF
                    wsMRS.Cells(outRow, 14).Formula = "=LunchFlag(" & rawDT & "," & rawDT & "+M" & r & "/1440,""12:00:00"",""00:00:00"",45)"
                Else
                    ' Static G: check both cross-lunch and during-lunch via G
                    wsMRS.Cells(outRow, 14).Formula = "=IF(OR(AND(G" & r & "<=TIME(12,0,0),G" & r & "+M" & r & "/1440>TIME(12,0,0)),AND(G" & r & ">=TIME(12,0,0),G" & r & "<=TIME(12,45,0))),""" & daStr & ""","""")"
                End If

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
            Dim dataEndRow As Long: dataEndRow = outRow
            outRow = outRow + 1
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 9)).Merge
            wsMRS.Cells(outRow, 2).Value = UW(1048, 1090, 1086, 1075, 1086) & ": " & blkCnt & " " & UW(1079, 1072, 1087, 1080, 1089, 1077, 1081)
            wsMRS.Cells(outRow, 2).Font.Bold = True
            wsMRS.Cells(outRow, 2).HorizontalAlignment = -4131
            wsMRS.Cells(outRow, 10).Formula = "=SUM(J" & dataStartRow & ":J" & dataEndRow & ")"
            wsMRS.Cells(outRow, 10).NumberFormat = "0.00"
            wsMRS.Cells(outRow, 10).Font.Bold = True
            wsMRS.Cells(outRow, 11).Formula = "=SUM(K" & dataStartRow & ":K" & dataEndRow & ")"
            wsMRS.Cells(outRow, 11).NumberFormat = "0.00"
            wsMRS.Cells(outRow, 11).Font.Bold = True
            wsMRS.Cells(outRow, 12).Formula = "=SUM(L" & dataStartRow & ":L" & dataEndRow & ")"
            wsMRS.Cells(outRow, 12).NumberFormat = "0.00"
            wsMRS.Cells(outRow, 12).Font.Bold = True
            wsMRS.Cells(outRow, 13).Formula = "=SUM(M" & dataStartRow & ":M" & dataEndRow & ")"
            wsMRS.Cells(outRow, 13).NumberFormat = "0.00"
            wsMRS.Cells(outRow, 13).Font.Bold = True
            wsMRS.Range(wsMRS.Cells(outRow, 2), wsMRS.Cells(outRow, 14)).Borders.LineStyle = xlContinuous
        Next oi

        outRow = outRow + 1 ' empty row between brigades
    Next b

    Next d

    ' Apply formatting to used area
    Dim dataRange As Range
    Set dataRange = wsMRS.Range(wsMRS.Cells(3, 2), wsMRS.Cells(outRow, 14))
    dataRange.Font.Size = 14
    dataRange.Font.Color = RGB(0, 0, 0)
    dataRange.HorizontalAlignment = -4108
    dataRange.VerticalAlignment = -4108

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
    errMsg = Err.Description
    Resume Done
End Sub

Public Sub ClearMRS()
    If Not ZQ() Then
        MsgBox UW(1054, 1096, 1080, 1073, 1082, 1072, 33), vbCritical
        Exit Sub
    End If
    On Error Resume Next
    Dim wsMRS As Worksheet
    Set wsMRS = ThisWorkbook.Worksheets(4)
    wsMRS.Unprotect UW(49, 49, 52, 55, 48, 57)
    ClearMRSArea wsMRS
    wsMRS.Protect UW(49, 49, 52, 55, 48, 57)
    On Error GoTo 0
End Sub

Private Sub ClearMRSArea(ByVal wsMRS As Worksheet)
    Dim rng As Range
    On Error Resume Next
    wsMRS.Range("B3:N20000").UnMerge
    On Error GoTo 0
    Set rng = wsMRS.Range("B3:N20000")
    rng.ClearContents
    rng.Interior.Color = RGB(255, 230, 230)
    rng.Borders.LineStyle = xlNone
    rng.HorizontalAlignment = -4108
    rng.VerticalAlignment = -4108
    rng.Font.Size = 14
    rng.Font.Color = RGB(0, 0, 0)
    rng.Font.Bold = False
    rng.NumberFormat = "General"
    rng.Locked = True
    rng.WrapText = False
    wsMRS.Rows("3:20000").RowHeight = wsMRS.StandardHeight
End Sub

