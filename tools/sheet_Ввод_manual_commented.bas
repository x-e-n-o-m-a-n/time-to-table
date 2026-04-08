' MIT License
'
' Copyright (c) 2026 Галимзянов Г.Р.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
' Автор исходного проекта: Галимзянов Г.Р.
' Примечание: данный файл представляет собой документированную версию sheet-модуля листа "Ввод".

Option Explicit

' ===========================================================================
' МОДУЛЬ ЛИСТА "Ввод"
' ===========================================================================
' Этот sheet-модуль управляет интерактивным поведением основного рабочего листа.
' Здесь сосредоточены события, которые:
'   - отслеживают изменение цветовых настроек;
'   - синхронизируют количество операций и исполнителей;
'   - нормализуют пользовательский ввод по датам, времени, числам и идентификаторам.
' ===========================================================================

' Хранит снимок текущих цветовых настроек листа
' для последующего сравнения между событиями.
Private mColorSignature As String
Private Const MAX_OPERATION_COUNT As Long = 20
Private Const FIRST_OPERATION_ROW As Long = 4
Private Const LAST_OPERATION_ROW As Long = FIRST_OPERATION_ROW + MAX_OPERATION_COUNT - 1
Private Const MAX_WORKER_COUNT As Long = 20

' ===========================================================================
' Worksheet_Activate
' ===========================================================================
' Назначение:
'   Сохраняет текущую сигнатуру цветовых настроек при активации листа.
Private Sub Worksheet_Activate()
    mColorSignature = BuildColorSettingsSignature(Me)
End Sub

' ===========================================================================
' Worksheet_SelectionChange
' ===========================================================================
' Назначение:
'   Проверяет, изменились ли пользовательские цвета при работе на листе,
'   и при необходимости обновляет цветовую схему книги.
' Параметры:
'   Target - новый выделенный диапазон.
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

' ===========================================================================
' Worksheet_Change
' ===========================================================================
' Назначение:
'   Выполняет основную валидацию, санитизацию и синхронизацию ввода
'   на главном рабочем листе.
' Параметры:
'   Target - изменённая ячейка или диапазон.
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    If Target Is Nothing Then Exit Sub

    Dim participantRange As Range, participantCell As Range

    ' На время пакетной обработки отключаем события и снимаем защиту листа.
    Application.EnableEvents = False
    Me.Unprotect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)

    Dim syncOpsFromSelectors As Boolean
    syncOpsFromSelectors = False

    ' Изменение числа операций сразу перестраивает видимый операционный блок.
    If Not Intersect(Target, Me.Range("B8")) Is Nothing Then
        Dim opCount As Long
        opCount = CLng(val(Me.Range("B8").Value))
        If opCount < 1 Then opCount = 1
        If opCount > MAX_OPERATION_COUNT Then opCount = MAX_OPERATION_COUNT
        Me.Range("B8").Value = opCount
        SyncOperationRows Me, opCount
    End If

    ' Изменение числа исполнителей синхронизирует блок их идентификаторов.
    If Not Intersect(Target, Me.Range("B9")) Is Nothing Then
        Dim workerCount As Long
        workerCount = CLng(val(Me.Range("B9").Value))
        If workerCount < 1 Then workerCount = 1
        If workerCount > MAX_WORKER_COUNT Then workerCount = MAX_WORKER_COUNT
        Me.Range("B9").Value = workerCount
        SyncWorkerIdInputs Me, workerCount
        Set participantRange = Me.Range("P" & FIRST_OPERATION_ROW & ":P" & LAST_OPERATION_ROW)
    End If

    ' Табельные номера исполнителей очищаются от лишних символов.
    Dim workerRange As Range, cell As Range
    Set workerRange = Intersect(Target, Me.Range("E4:E23"))
    If Not workerRange Is Nothing Then
        For Each cell In workerRange.Cells
            SanitizeWorkerIdCell cell
        Next cell
    End If

    ' Номер заказа в B3 приводится к фиксированному цифровому формату.
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

    ' Поле ПДТВ ограничивается восемью цифрами.
    Dim hRange As Range, hCell As Range
    Set hRange = Intersect(Target, Me.Range("H" & FIRST_OPERATION_ROW & ":H" & LAST_OPERATION_ROW))
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

    ' Текстовые десятичные значения приводятся к единому разделителю.
    Dim jRange As Range, jCell As Range
    Set jRange = Intersect(Target, Me.Range("J" & FIRST_OPERATION_ROW & ":J" & LAST_OPERATION_ROW))
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

    ' Числовые поля длительностей и пауз очищаются и нормализуются.
    Dim kmRange As Range, kmCell As Range
    Set kmRange = Intersect(Target, Me.Range("K" & FIRST_OPERATION_ROW & ":K" & LAST_OPERATION_ROW & ",N" & FIRST_OPERATION_ROW & ":N" & LAST_OPERATION_ROW))
    If Not kmRange Is Nothing Then
        For Each kmCell In kmRange.Cells
            Dim cleanedKM As String
            cleanedKM = NormalizeDecimal(CStr(kmCell.Value))
            If cleanedKM = "" Then
                kmCell.Value = 0
            Else
                kmCell.Value = val(Replace(cleanedKM, ",", "."))
            End If
        Next kmCell
    End If

    ' Специализированные параметры проходят через собственные нормализаторы.
    If Not Intersect(Target, Me.Range("B16")) Is Nothing Then
        Me.Range("B16").Value = NormalizeRizInput(CStr(Me.Range("B16").Value))
    End If

    If Not Intersect(Target, Me.Range("B17")) Is Nothing Then
        Me.Range("B17").Value = NormalizeKInput(CStr(Me.Range("B17").Value))
    End If

    ' Временные поля приводятся к корректному формату времени.
    Dim inputTimeRange As Range, inputTimeCell As Range
    Set inputTimeRange = Intersect(Target, Union(Me.Range("B6"), Me.Range("B10:B11")))
    If Not inputTimeRange Is Nothing Then
        For Each inputTimeCell In inputTimeRange.Cells
            SanitizeTimeCell inputTimeCell
        Next inputTimeCell
    End If

    ' Участники операций могут быть затронуты напрямую или косвенно.
    If participantRange Is Nothing Then
        Set participantRange = Intersect(Target, Me.Range("P" & FIRST_OPERATION_ROW & ":P" & LAST_OPERATION_ROW))
    ElseIf Not Intersect(Target, Me.Range("P" & FIRST_OPERATION_ROW & ":P" & LAST_OPERATION_ROW)) Is Nothing Then
        Set participantRange = Union(participantRange, Intersect(Target, Me.Range("P" & FIRST_OPERATION_ROW & ":P" & LAST_OPERATION_ROW)))
    End If

    If Not participantRange Is Nothing Then
        Dim maxParticipantWorkerCount As Long
        maxParticipantWorkerCount = CLng(val(Me.Range("B9").Value))
        If maxParticipantWorkerCount < 1 Then maxParticipantWorkerCount = 1
        If maxParticipantWorkerCount > MAX_WORKER_COUNT Then maxParticipantWorkerCount = MAX_WORKER_COUNT

        participantRange.NumberFormat = "@"
        For Each participantCell In participantRange.Cells
            SanitizeParticipantsCell participantCell, maxParticipantWorkerCount
        Next participantCell
    End If

    ' Изменение селекторов режима расчёта требует синхронизации строк операций.
    If Not Intersect(Target, Me.Range("L" & FIRST_OPERATION_ROW)) Is Nothing Then
        syncOpsFromSelectors = True
    End If

    If Not Intersect(Target, Me.Range("M" & FIRST_OPERATION_ROW)) Is Nothing Then
        syncOpsFromSelectors = True
    End If

    If Not Intersect(Target, Me.Range("O" & FIRST_OPERATION_ROW)) Is Nothing Then
        syncOpsFromSelectors = True
    End If

    If syncOpsFromSelectors Then
        Dim currentOpCount As Long
        currentOpCount = CLng(val(Me.Range("B8").Value))
        If currentOpCount < 1 Then currentOpCount = 1
        If currentOpCount > MAX_OPERATION_COUNT Then currentOpCount = MAX_OPERATION_COUNT
        SyncOperationRows Me, currentOpCount
    End If

SafeExit:
    ' Возвращаем сигнатуру, защиту листа и обычный режим событий.
    mColorSignature = BuildColorSettingsSignature(Me)
    Me.Protect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)
    Application.EnableEvents = True
End Sub
