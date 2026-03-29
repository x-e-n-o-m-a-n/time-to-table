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
' Примечание: данный файл представляет собой документированную версию модуля ThisWorkbook.

' ===========================================================================
' МОДУЛЬ ThisWorkbook
' ===========================================================================
' Модуль книги содержит глобальные обработчики событий workbook-уровня.
' Здесь размещаются сценарии, которые должны запускаться автоматически:
'   - при открытии книги;
'   - при изменении данных на отдельных листах;
'   - перед сохранением книги пользователем.
' ===========================================================================

' ===========================================================================
' Workbook_Open
' ===========================================================================
' Назначение:
'   Выполняет начальную инициализацию книги после открытия.
' Основные действия:
'   - получает ссылки на стартовые листы;
'   - обновляет текущую дату в ключевых полях листа ввода;
'   - запускает синхронизацию цветовой схемы;
'   - переводит пользователя на стартовый лист.
Private Sub Workbook_Open()
    Dim wsDisclaimer As Worksheet
    Dim wsIn As Worksheet
    Set wsDisclaimer = ThisWorkbook.Worksheets(1)
    Set wsIn = ThisWorkbook.Worksheets(2)

    ' Инициализируем рабочие даты на листе ввода.
    wsIn.Unprotect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)
    wsIn.Range("B5").Value = Date
    wsIn.Range("B7").Value = Date
    wsIn.Protect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)

    ' После обновления данных приводим оформление к актуальному состоянию.
    RefreshWorkbookColors
    wsDisclaimer.Activate
    wsDisclaimer.Range("B2").Select
End Sub

' ===========================================================================
' Workbook_SheetChange
' ===========================================================================
' Назначение:
'   Централизованно перенаправляет изменения на листах книги
'   в специальные обработчики общего модуля.
' Параметры:
'   Sh - лист, на котором произошло изменение.
'   Target - изменённая ячейка или диапазон.
' Эффект выполнения:
'   Листы "История" и "Парсинг MRS" получают свою логику через dispatcher книги,
'   а не через собственные sheet-level обработчики.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Target Is Nothing Then Exit Sub
    If Not TypeOf Sh Is Worksheet Then Exit Sub

    If Sh Is ThisWorkbook.Worksheets(4) Then
        HandleHistorySheetChange Sh, Target
    ElseIf Sh Is ThisWorkbook.Worksheets(5) Then
        HandleMRSSheetChange Sh, Target
    End If
End Sub

' ===========================================================================
' Workbook_BeforeSave
' ===========================================================================
' Назначение:
'   Выполняет пользовательскую проверку перед сохранением книги.
' Основные действия:
'   - формирует текст приглашения;
'   - запрашивает подтверждающую строку;
'   - при неуспешной проверке отменяет сохранение.
' Параметры:
'   SaveAsUI - системный флаг Excel для сценария Save As.
'   Cancel - флаг отмены сохранения.
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim userInput As String
    Dim correctPassword As String
    Dim msg As String
    Dim promptText As String
    Dim titleText As String

    correctPassword = UW(49, 48, 53, 51, 53, 53)
    
    promptText = UW(1044, 1083, 1103, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103, 32, 1092, 1072, 1081, 1083, 1072, 32, 1090, 1088, 1077, 1073, 1091, 1077, 1090, 1089, 1103, 32, 1087, 1072, 1088, 1086, 1083, 1100, 32, 1072, 1076, 1084, 1080, 1085, 1080, 1089, 1090, 1088, 1072, 1090, 1086, 1088, 1072, 46) _
                 & vbCrLf & vbCrLf & _
                 UW(1042, 1074, 1077, 1076, 1080, 1090, 1077, 32, 1087, 1072, 1088, 1086, 1083, 1100, 58)
    
    titleText = UW(1055, 1086, 1076, 1090, 1074, 1077, 1088, 1078, 1076, 1077, 1085, 1080, 1077, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1077, 1085, 1080, 1103)
    
    ' Запрашиваем подтверждение перед сохранением книги.
    userInput = InputBox(promptText, titleText)
    
    If userInput = correctPassword Then
    Else
        msg = UW(1058, 1072, 1073, 1077, 1083, 1100, 1085, 1099, 1077, 32, 1085, 1086, 1084, 1077, 1088, 1072, 32, 1087, 1086, 1076, 1087, 1072, 1076, 1072, 1102, 1090, 32, 1087, 1086, 1076, 32, 1079, 1072, 1082, 1086, 1085, 32, 1086, 1073, 32, 1086, 1073, 1088, 1072, 1073, 1086, 1090, 1082, 1077, 32, 1087, 1077, 1088, 1089, 1086, 1085, 1072, 1083, 1100, 1085, 1099, 1093, 32, 1076, 1072, 1085, 1085, 1099, 1093, 44, 32, 1089, 1086, 1093, 1088, 1072, 1085, 1103, 1090, 1100, 47, 1088, 1077, 1076, 1072, 1082, 1090, 1080, 1088, 1086, 1074, 1072, 1090, 1100, 32, 1096, 1072, 1073, 1083, 1086, 1085, 32, 1079, 1072, 1087, 1088, 1077, 1097, 1077, 1085, 1086, 46)
        
        MsgBox msg, vbCritical, UW(1044, 1077, 1081, 1089, 1090, 1074, 1080, 1077, 32, 1079, 1072, 1087, 1088, 1077, 1097, 1077, 1085, 1086)
        Cancel = True
    End If
End Sub

