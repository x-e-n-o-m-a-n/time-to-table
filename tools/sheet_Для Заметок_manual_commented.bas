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
' Примечание: данный файл представляет собой документированную версию sheet-модуля листа "Для Заметок".

Option Explicit

' ===========================================================================
' МОДУЛЬ ЛИСТА "Для Заметок"
' ===========================================================================
' Лист заметок содержит небольшой обработчик ввода.
' Его задача сводится к нормализации табельных номеров
' в рабочем диапазоне при ручном редактировании.
' ===========================================================================

' ===========================================================================
' Worksheet_Change
' ===========================================================================
' Назначение:
'   Обрабатывает изменения на листе заметок в диапазоне табельных номеров.
' Параметры:
'   Target - изменённая ячейка или диапазон.
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    If Target Is Nothing Then Exit Sub

    Dim workerRange As Range
    Set workerRange = Intersect(Target, Me.Range("B3:B52"))
    If workerRange Is Nothing Then Exit Sub

    ' На время обработки отключаем события и снимаем защиту листа.
    Application.EnableEvents = False
    Me.Unprotect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)

    Dim cell As Range
    For Each cell In workerRange.Cells
        SanitizeWorkerIdCell cell
    Next cell

SafeExit:
    Me.Protect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)
    Application.EnableEvents = True
End Sub

