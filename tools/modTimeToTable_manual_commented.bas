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
' Примечание: данный файл представляет собой документированную версию VBA-модуля.

Option Explicit

' ===========================================================================
' МОДУЛЬ modTimeToTable
' ===========================================================================
' Модуль объединяет несколько ключевых подсистем книги:
'   - хранение и применение пользовательских цветовых настроек;
'   - стилизацию и защиту рабочих листов;
'   - нормализацию и валидацию пользовательского ввода;
'   - расчёт цепочки операций с учётом исполнителей, пауз и обеда;
'   - перенос результата на лист истории;
'   - импорт внешних данных MRS и их группировку;
'   - экспорт листов истории и MRS в отдельные xlsx-файлы.
' ===========================================================================


' ---------------------------------------------------------------------------
' КОНСТАНТЫ: строки и колонки блока цветовых настроек
' ---------------------------------------------------------------------------
' В этой группе задаются координаты блока настроек цветов на листе ввода.
' Эти значения используются при чтении, отображении и восстановлении цветов интерфейса.
Private Const CLR_ROW_LOCKED As Long = 31
Private Const CLR_ROW_EDITABLE As Long = 32
Private Const CLR_ROW_MRS_HEADER As Long = 33
Private Const CLR_ROW_MRS_SUBHEADER As Long = 34
Private Const CLR_ROW_MRS_ORDER As Long = 35
Private Const CLR_ROW_MRS_ORDER_UNCONF As Long = 36
Private Const CLR_ROW_HEADER As Long = 37
Private Const CLR_COL As Long = 2


' ---------------------------------------------------------------------------
' КОНСТАНТЫ: значения цветов по умолчанию
' ---------------------------------------------------------------------------
' Значения из этой группы применяются, когда пользователь ещё не задал
' собственные цвета или когда соответствующие ячейки настроек пусты.
Private Const CLR_DEF_LOCKED As Long = 15132415
Private Const CLR_DEF_MRS_HEADER As Long = 15189684
Private Const CLR_DEF_MRS_SUB As Long = 16768200
Private Const CLR_DEF_MRS_ORDER As Long = 13167560
Private Const CLR_DEF_MRS_ORDER_UNCONF As Long = 14277081
Private Const CLR_DEF_HEADER As Long = 13167560

' ===========================================================================
' ReadCellColor
' ===========================================================================
' Назначение:
'   Считывает реальный цвет заливки из ячейки Excel.
'   Если ячейка визуально не окрашена или Excel возвращает нулевое значение
'   цвета, функция подставляет заранее заданный цвет по умолчанию.
' Основные действия:
'   - Проверяет наличие паттерна заливки.
'   - Учитывает случай, когда цвет формально отсутствует.
'   - Возвращает пригодное для дальнейшей стилизации значение Long.
' Параметры:
'   cell - ячейка, из которой считывается цвет.
'   defaultColor - резервный цвет, используемый при отсутствии явной заливки.
' Возвращаемое значение:
'   Long-код цвета, который можно безопасно использовать в оформлении листов.
Private Function ReadCellColor(ByVal cell As Range, ByVal defaultColor As Long) As Long
    If cell.Interior.Pattern = xlNone Or cell.Interior.Color = 0 Then
        ReadCellColor = defaultColor
    Else
        ReadCellColor = cell.Interior.Color
    End If
End Function

' ===========================================================================
' GetContrastColor
' ===========================================================================
' Назначение:
'   Подбирает цвет текста, который будет читаться на заданном фоне.
'   Используется везде, где строки динамически перекрашиваются по статусу или
'   пользовательским цветовым настройкам.
' Основные действия:
'   - Разбирает цвет Excel в формате Long на компоненты RGB.
'   - Вычисляет условную яркость фона.
'   - Возвращает чёрный или белый цвет в зависимости от яркости.
' Параметры:
'   bgColor - цвет фона, для которого выбирается контрастный шрифт.
' Возвращаемое значение:
'   Long-код цвета шрифта: обычно чёрный для светлого фона и белый для тёмного.
Public Function GetContrastColor(ByVal bgColor As Long) As Long
    If bgColor = xlNone Or bgColor = -4142 Then
        GetContrastColor = 0
        Exit Function
    End If
    
    Dim r As Long, G As Long, b As Long
    Dim luminance As Double
    
    r = bgColor Mod 256
    G = (bgColor \ 256) Mod 256
    b = (bgColor \ 65536) Mod 256
    
    luminance = (r * 299& + G * 587& + b * 114&) / 1000#
    
    If luminance > 128 Then
        GetContrastColor = 0
    Else
        GetContrastColor = 16777215
    End If
End Function

' ===========================================================================
' ReadAllColors
' ===========================================================================
' Назначение:
'   Читает полный набор активных цветовых настроек с листа "Ввод".
'   Это центральная точка получения палитры, которой затем пользуются все
'   процедуры перекраски интерфейса и строк данных.
' Основные действия:
'   - Считывает цвет заблокированных ячеек.
'   - Определяет, задан ли отдельный цвет для редактируемых ячеек.
'   - Получает цвета заголовков MRS, подтверждённых и неподтверждённых заказов,
'     а также цвет шапок и служебных областей.
'   - Возвращает результаты через параметры ByRef.
' Параметры:
'   wsIn - лист "Ввод", где расположен блок настройки цветов.
'   clrLocked - сюда возвращается цвет заблокированных ячеек.
'   clrEditHasColor - флаг, указывающий, задан ли отдельный цвет редактируемых ячеек.
'   clrEditable - сюда возвращается цвет редактируемых ячеек.
'   clrMrsHeader - сюда возвращается цвет заголовков дат MRS.
'   clrMrsSub - сюда возвращается цвет строк бригад MRS.
'   clrMrsOrder - сюда возвращается цвет подтверждённых заказов.
'   clrMrsOrderUnconf - сюда возвращается цвет неподтверждённых заказов.
'   clrHeader - сюда возвращается цвет шапок и выделенных служебных областей.
' Эффект выполнения:
'   Процедура не меняет лист напрямую, а заполняет набор переменных для других сценариев.
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

' ===========================================================================
' GetOrderColors
' ===========================================================================
' Назначение:
'   Быстро возвращает цвета, связанные со строками заказов.
'   Процедура нужна в обработчиках изменений листов и при экспорте, где нет
'   необходимости заново читать весь набор пользовательской палитры.
' Основные действия:
'   - Берёт лист "Ввод" как источник настроек.
'   - Читает цвет подтверждённых заказов.
'   - Читает цвет неподтверждённых заказов.
'   - При необходимости дополнительно возвращает цвет блокировки и состояние
'     оформления редактируемых ячеек.
' Параметры:
'   clrMrsOrder - сюда возвращается цвет подтверждённых заказов.
'   clrMrsOrderUnconf - сюда возвращается цвет неподтверждённых заказов.
'   clrLocked - необязательный выходной параметр для цвета блокировки.
'   clrEditHasColor - необязательный флаг наличия цвета редактируемых ячеек.
'   clrEditable - необязательный выходной параметр для цвета редактируемых ячеек.
' Эффект выполнения:
'   Процедура читает настройки из книги и возвращает их вызывающему коду.
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

' ===========================================================================
' ColorOrderRow
' ===========================================================================
' Назначение:
'   Применяет цветовое оформление к целой строке заказа по значению статуса в колонке B.
'   Используется как на листе истории, так и на листе MRS.
' Основные действия:
'   - Берёт диапазон строки от колонки B до последней рабочей колонки.
'   - Считывает текст статуса в колонке B.
'   - Если заказ подтверждён, применяет основной цвет заказа.
'   - Если заказ не подтверждён, применяет альтернативный цвет.
'   - Для текста автоматически подбирает контрастный цвет шрифта.
' Параметры:
'   ws - лист, на котором находится строка заказа.
'   r - номер окрашиваемой строки.
'   lastCol - последняя колонка строки, входящая в цветовой диапазон.
'   clrOrder - цвет подтверждённого заказа.
'   clrUnconf - цвет неподтверждённого заказа.
' Эффект выполнения:
'   Процедура меняет заливку и цвет шрифта в строке заказа.
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

' ===========================================================================
' EnsureColorSettings
' ===========================================================================
' Назначение:
'   Восстанавливает и поддерживает служебный блок настройки цветов на листе "Ввод".
'   Без этого блока пользователь не сможет менять палитру, а процедуры перекраски
'   книги не получат корректный источник настроек.
' Основные действия:
'   - Создаёт заголовок блока и подписи строк настроек.
'   - Делает текстовые метки защищёнными, а ячейки выбора цвета редактируемыми.
'   - Заполняет пустые цветовые настройки цветами по умолчанию.
'   - Временно отключает события, чтобы не вызвать рекурсивные обработчики.
' Параметры:
'   wsIn - лист "Ввод", на котором хранится блок цветовых настроек.
' Эффект выполнения:
'   Процедура модифицирует разметку, защиту и цветовые образцы на листе ввода.
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
    ' Объединение заголовка может уже существовать, поэтому merge выполняется
    ' через отдельный quiet-helper без переключения внешнего обработчика ошибок.
    MergeRangeQuiet wsIn.Range("A29:B29")
    wsIn.Cells(29, 1).Font.Bold = True

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

' ===========================================================================
' ApplyLockedStyle
' ===========================================================================
' Назначение:
'   Применяет к диапазону стиль "заблокированная область".
'   Такой стиль используется для ячеек, которые пользователь не должен менять напрямую.
' Основные действия:
'   - Задаёт диапазону цвет фона заблокированной области.
'   - Подбирает для него контрастный цвет текста.
' Параметры:
'   rng - диапазон, который нужно оформить.
'   clrLocked - цвет заливки для заблокированного состояния.
' Эффект выполнения:
'   Процедура меняет визуальное оформление переданного диапазона.
Private Sub ApplyLockedStyle(ByVal rng As Range, ByVal clrLocked As Long)
    rng.Interior.Color = clrLocked
    rng.Font.Color = GetContrastColor(clrLocked)
End Sub

' ===========================================================================
' ApplyEditableStyle
' ===========================================================================
' Назначение:
'   Применяет к диапазону стиль редактируемой области.
'   Поддерживает два режима: явная заливка для редактируемых ячеек и прозрачный
'   режим, когда пользователь оставил редактируемые ячейки без фона.
' Основные действия:
'   - Если пользовательский цвет задан, применяет его.
'   - Если цвет не задан, убирает заливку.
'   - Настраивает цвет шрифта под выбранное состояние.
' Параметры:
'   rng - диапазон, который нужно оформить как редактируемый.
'   clrEditHasColor - флаг, показывающий, задан ли цвет редактируемых ячеек.
'   clrEditable - цвет редактируемых ячеек.
' Эффект выполнения:
'   Процедура меняет заливку и цвет текста переданного диапазона.
Private Sub ApplyEditableStyle(ByVal rng As Range, ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long)
    If clrEditHasColor Then
        rng.Interior.Color = clrEditable
        rng.Font.Color = GetContrastColor(clrEditable)
    Else
        rng.Interior.Pattern = xlNone
        rng.Font.Color = 0
    End If
End Sub

' ===========================================================================
' BuildColorSettingsSignature
' ===========================================================================
' Назначение:
'   Строит компактную строковую сигнатуру блока цветовых настроек.
'   Такая сигнатура удобна для сравнения: если строка изменилась, значит пользователь
'   поменял один из цветов или состояние соответствующей ячейки.
' Основные действия:
'   - Проходит по строкам служебного блока настроек цветов.
'   - Для каждой строки фиксирует паттерн заливки и сам цвет.
'   - Собирает результаты в одну строку с разделителями.
' Параметры:
'   wsIn - лист "Ввод", содержащий блок цветовых настроек.
' Возвращаемое значение:
'   Строка-сигнатура, описывающая текущее состояние цветового блока.
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

' ===========================================================================
' LastContentRow
' ===========================================================================
' Назначение:
'   Находит нижнюю границу данных на листе.
'   Это базовый helper для построения динамических диапазонов без жёстко заданной длины.
' Основные действия:
'   - Ищет последнюю непустую ячейку методом `Find`.
'   - Если данных нет, возвращает безопасное минимальное значение.
' Параметры:
'   ws - лист, на котором ищется последняя заполненная строка.
'   minRow - минимально допустимое значение результата.
' Возвращаемое значение:
'   Номер последней содержательной строки, но не меньше `minRow`.
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

' ===========================================================================
' ApplyHistoryOperationSeparators
' ===========================================================================
' Назначение:
'   Ставит визуальные разделители между операциями на листе истории.
'   Это упрощает чтение больших блоков истории, где подряд могут идти строки
'   нескольких операций и нескольких исполнителей.
' Основные действия:
'   - Сначала назначает всем строкам стандартную нижнюю границу.
'   - Затем ищет места, где номер операции меняется.
'   - На границе смены операции ставит утолщённую линию.
' Параметры:
'   wsHist - лист истории.
' Эффект выполнения:
'   На листе истории появляются визуальные разделители между соседними операциями.
Private Sub ApplyHistoryOperationSeparators(ByVal wsHist As Worksheet)
    Dim lastRow As Long, r As Long
    Dim currOp As String, nextOp As String

    lastRow = LastContentRow(wsHist, 4)
    If lastRow < 4 Then Exit Sub

    For r = 4 To lastRow
        With wsHist.Range(wsHist.Cells(r, 2), wsHist.Cells(r, 22)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next r

    For r = 4 To lastRow - 1
        If IsNumeric(wsHist.Cells(r, 2).Value) And IsNumeric(wsHist.Cells(r + 1, 2).Value) Then
            currOp = Trim$(CStr(wsHist.Cells(r, 2).Value))
            nextOp = Trim$(CStr(wsHist.Cells(r + 1, 2).Value))
            If Len(currOp) > 0 And Len(nextOp) > 0 And currOp <> nextOp Then
                With wsHist.Range(wsHist.Cells(r, 2), wsHist.Cells(r, 22)).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThick
                End With
            End If
        End If
    Next r
End Sub

' ===========================================================================
' IsMRSDataRow
' ===========================================================================
' Назначение:
'   Определяет, является ли строка листа MRS строкой данных, а не заголовком.
'   Нужна для корректной перекраски и постановки разделителей только на строках операций.
' Основные действия:
'   - Отбрасывает строки шапки.
'   - Исключает объединённые заголовочные строки.
'   - Проверяет наличие числового номера и названия операции.
' Параметры:
'   wsMRS - лист MRS.
'   rowNum - номер проверяемой строки.
' Возвращаемое значение:
'   True, если строка содержит данные операции MRS; иначе False.
Private Function IsMRSDataRow(ByVal wsMRS As Worksheet, ByVal rowNum As Long) As Boolean
    If rowNum < 4 Then Exit Function
    If wsMRS.Cells(rowNum, 2).MergeCells Then Exit Function
    If Not IsNumeric(wsMRS.Cells(rowNum, 2).Value) Then Exit Function
    If Len(Trim$(CStr(wsMRS.Cells(rowNum, 3).Value))) = 0 Then Exit Function
    IsMRSDataRow = True
End Function

' ===========================================================================
' ApplyMRSOperationSeparators
' ===========================================================================
' Назначение:
'   Ставит визуальные разделители между разными операциями на листе MRS.
'   Аналогична процедуре для истории, но работает по структуре и ширине листа MRS.
' Основные действия:
'   - Назначает всем строкам данных обычную нижнюю границу.
'   - Сравнивает название текущей и следующей операции.
'   - При смене операции делает границу утолщённой.
' Параметры:
'   wsMRS - лист парсинга MRS.
' Эффект выполнения:
'   Группы операций на листе MRS становятся визуально отделёнными.
Private Sub ApplyMRSOperationSeparators(ByVal wsMRS As Worksheet)
    Dim lastRow As Long, r As Long
    Dim currOp As String, nextOp As String

    lastRow = LastContentRow(wsMRS, 4)
    If lastRow < 4 Then Exit Sub

    For r = 4 To lastRow
        With wsMRS.Range(wsMRS.Cells(r, 2), wsMRS.Cells(r, 14)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next r

    For r = 4 To lastRow - 1
        If IsMRSDataRow(wsMRS, r) And IsMRSDataRow(wsMRS, r + 1) Then
            currOp = Trim$(CStr(wsMRS.Cells(r, 3).Value))
            nextOp = Trim$(CStr(wsMRS.Cells(r + 1, 3).Value))
            If Len(currOp) > 0 And Len(nextOp) > 0 And StrComp(currOp, nextOp, vbTextCompare) <> 0 Then
                With wsMRS.Range(wsMRS.Cells(r, 2), wsMRS.Cells(r, 14)).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThick
                End With
            End If
        End If
    Next r
End Sub

' ===========================================================================
' SyncPauseInputCell
' ===========================================================================
' Назначение:
'   Синхронизирует доступность ячейки паузы перед первой операцией.
'   Пока история пуста, такая пауза не имеет смысла, поэтому поле должно быть
'   недоступным; после появления истории поле становится редактируемым.
' Основные действия:
'   - Проверяет, есть ли на листе истории строки данных.
'   - Если истории нет, блокирует ячейку паузы и оформляет её как служебную.
'   - Если история есть, делает ячейку редактируемой.
'   - В любом случае восстанавливает границу ячейки.
' Параметры:
'   wsIn - лист ввода.
'   wsHist - лист истории.
'   clrLocked - цвет заблокированного состояния.
'   clrEditHasColor - признак пользовательского цвета редактируемых ячеек.
'   clrEditable - цвет редактируемой ячейки.
' Эффект выполнения:
'   Поле паузы на листе ввода всегда соответствует текущему состоянию истории.
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

' ===========================================================================
' RefreshInputSheetColors
' ===========================================================================
' Назначение:
'   Полностью обновляет цветовую схему листа "Ввод".
'   Это наиболее насыщенный по структуре лист книги, поэтому для него отдельная
'   процедура перекраски нужна и для базовых полей, и для динамических блоков.
' Основные действия:
'   - Применяет общий фон заблокированного состояния ко всему листу.
'   - Выделяет все редактируемые области.
'   - Восстанавливает образцы цветовых настроек в служебном блоке.
'   - Синхронизирует количество строк работников и операций.
'   - Обновляет состояние ячейки паузы перед первой операцией.
' Параметры:
'   wsIn - лист ввода.
'   wsHist - лист истории.
'   clrLocked - базовый цвет заблокированных областей.
'   clrEditHasColor - признак пользовательского цвета редактируемых ячеек.
'   clrEditable - цвет редактируемых ячеек.
'   clrMrsHeader - цвет заголовков дат MRS.
'   clrMrsSub - цвет строк бригад MRS.
'   clrMrsOrder - цвет подтверждённых заказов.
'   clrMrsOrderUnconf - цвет неподтверждённых заказов.
'   clrHeader - цвет шапок и служебных областей.
' Эффект выполнения:
'   Лист ввода полностью приводится к актуальной палитре и корректной структуре интерфейса.
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

    workerCount = CLng(val(wsIn.Range("B9").Value))
    If workerCount < 1 Then workerCount = 1
    If workerCount > 10 Then workerCount = 10
    SyncWorkerIdInputs wsIn, workerCount

    opCount = CLng(val(wsIn.Range("B8").Value))
    If opCount < 1 Then opCount = 1
    If opCount > 20 Then opCount = 20
    SyncOperationRows wsIn, opCount

    SyncPauseInputCell wsIn, wsHist, clrLocked, clrEditHasColor, clrEditable
End Sub

' ===========================================================================
' RefreshResultSheetColors
' ===========================================================================
' Назначение:
'   Обновляет базовое оформление листа результата.
'   Лист результата почти целиком служебный, поэтому его перекраска сводится
'   к единому фону и контрастному цвету текста.
' Параметры:
'   wsOut - лист результата.
'   clrLocked - базовый цвет служебного состояния.
' Эффект выполнения:
'   Лист результата получает единый фон и согласованный цвет шрифта.
Private Sub RefreshResultSheetColors(ByVal wsOut As Worksheet, ByVal clrLocked As Long)
    wsOut.Cells.Interior.Color = clrLocked
    wsOut.Cells.Font.Color = GetContrastColor(clrLocked)
End Sub

' ===========================================================================
' RefreshDisclaimerSheetColors
' ===========================================================================
' Назначение:
'   Обновляет цветовую схему служебного листа дисклеймера.
'   Лист не содержит сложной интерактивной структуры, поэтому для него достаточно
'   единого фона и контрастного текста.
' Параметры:
'   wsDisclaimer - лист дисклеймера.
'   clrLocked - базовый цвет служебного фона.
' Эффект выполнения:
'   Лист дисклеймера приводится к общей палитре книги.
Private Sub RefreshDisclaimerSheetColors(ByVal wsDisclaimer As Worksheet, ByVal clrLocked As Long)
    wsDisclaimer.Cells.Interior.Color = clrLocked
    wsDisclaimer.Cells.Font.Color = GetContrastColor(clrLocked)
End Sub

' ===========================================================================
' RefreshHistorySheetColors
' ===========================================================================
' Назначение:
'   Полностью перекрашивает лист истории.
'   Процедура должна учесть и служебные шапки, и строки заказов по статусу,
'   и отдельные редактируемые колонки внутри блока истории.
' Основные действия:
'   - Заливает весь лист базовым фоном.
'   - Перекрашивает верхнюю шапку и служебные строки.
'   - Разносит строки на подтверждённые и неподтверждённые заказы.
'   - Возвращает стиль редактируемым колонкам.
'   - Ставит визуальные разделители операций.
' Параметры:
'   wsHist - лист истории.
'   clrLocked - цвет базового фона.
'   clrEditHasColor - признак пользовательского цвета редактируемых ячеек.
'   clrEditable - цвет редактируемых ячеек.
'   clrMrsOrder - цвет подтверждённых заказов.
'   clrMrsOrderUnconf - цвет неподтверждённых заказов.
'   clrHeader - цвет верхних шапок и служебных зон.
' Эффект выполнения:
'   История получает актуальную палитру и остаётся читаемой даже после массовых изменений.
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

    editCols = Array(5, 7, 12, 15, 16, 18, 19)
    For idx = LBound(editCols) To UBound(editCols)
        For r = 2 To lastRow
            If Not wsHist.Cells(r, editCols(idx)).Locked Then
                ApplyEditableStyle wsHist.Cells(r, editCols(idx)), clrEditHasColor, clrEditable
            End If
        Next r
    Next idx

    ApplyHistoryOperationSeparators wsHist
End Sub

' ===========================================================================
' RefreshMRSSheetColors
' ===========================================================================
' Назначение:
'   Полностью перекрашивает лист "Парсинг MRS".
'   Лист содержит несколько типов строк: шапки, заголовки дат, заголовки бригад
'   и собственно строки заказов, поэтому оформление разбирается по типам.
' Основные действия:
'   - Применяет базовый фон ко всему листу.
'   - Отдельно перекрашивает верхние строки-шапки.
'   - Собирает диапазоны строк разных типов и красит их по своей палитре.
'   - Возвращает стиль редактируемым колонкам.
'   - Восстанавливает визуальные разделители операций.
' Параметры:
'   wsMRS - лист парсинга MRS.
'   clrLocked - цвет базового фона.
'   clrEditHasColor - признак пользовательского цвета редактируемых ячеек.
'   clrEditable - цвет редактируемых ячеек.
'   clrMrsHeader - цвет заголовков дат.
'   clrMrsSub - цвет строк бригад.
'   clrMrsOrder - цвет подтверждённых заказов.
'   clrMrsOrderUnconf - цвет неподтверждённых заказов.
'   clrHeader - цвет верхней шапки листа.
' Эффект выполнения:
'   Лист MRS приводится к согласованному виду после импорта, очистки или смены палитры.
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

    ApplyMRSOperationSeparators wsMRS
End Sub

' ===========================================================================
' RefreshNotesSheetColors
' ===========================================================================
' Назначение:
'   Обновляет цветовую схему листа заметок.
'   Лист заметок имеет собственную компактную структуру, поэтому для него используется
'   отдельная процедура перекраски, а не общая логика рабочих листов.
' Основные действия:
'   - Заливает весь лист базовым цветом заблокированного состояния.
'   - Отдельно оформляет шапки таблиц заметок.
'   - Делает редактируемые колонки заметок визуально отличимыми.
'   - Восстанавливает границы таблиц.
' Параметры:
'   wsNotes - лист заметок.
'   clrLocked - базовый цвет заблокированных областей.
'   clrEditHasColor - флаг наличия цвета для редактируемых ячеек.
'   clrEditable - цвет редактируемых полей.
'   clrHeader - цвет заголовков таблиц заметок.
' Эффект выполнения:
'   Лист заметок приводится к актуальной цветовой схеме книги.
Private Sub RefreshNotesSheetColors(ByVal wsNotes As Worksheet, ByVal clrLocked As Long, _
    ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long, ByVal clrHeader As Long)

    wsNotes.Cells.Interior.Color = clrLocked
    wsNotes.Cells.Font.Color = GetContrastColor(clrLocked)

    wsNotes.Range("B2:C2").Interior.Color = clrHeader
    wsNotes.Range("B2:C2").Font.Color = GetContrastColor(clrHeader)
    wsNotes.Range("F2").Interior.Color = clrHeader
    wsNotes.Range("F2").Font.Color = GetContrastColor(clrHeader)

    ApplyEditableStyle wsNotes.Range("B3:B52"), clrEditHasColor, clrEditable
    ApplyEditableStyle wsNotes.Range("C3:C52"), clrEditHasColor, clrEditable
    ApplyEditableStyle wsNotes.Range("F3:F52"), clrEditHasColor, clrEditable

    wsNotes.Range("B2:C52").Borders.LineStyle = xlContinuous
    wsNotes.Range("F2:F52").Borders.LineStyle = xlContinuous
End Sub

' ===========================================================================
' SetRangeBoldSafe
' ===========================================================================
' Назначение:
'   Безопасно устанавливает или снимает жирное начертание для диапазона.
'   В отличие от прямого `Range.Font.Bold = ...`, процедура аккуратно обходит
'   объединённые ячейки и поднимает более информативную ошибку при сбое.
' Основные действия:
'   - Проходит по всем ячейкам диапазона.
'   - Для объединённых областей меняет стиль только у первой ячейки MergeArea.
'   - При ошибке возвращает расширенное описание с именем этапа и адресом ячейки.
' Параметры:
'   Target - диапазон, где нужно изменить жирность текста.
'   isBold - признак, нужно ли включить жирное начертание.
'   tag - служебная текстовая метка для диагностики.
' Эффект выполнения:
'   Процедура меняет форматирование текста в диапазоне и упрощает поиск ошибок форматирования.
Private Sub SetRangeBoldSafe(ByVal Target As Range, ByVal isBold As Boolean, Optional ByVal tag As String = "")
    Dim cell As Range

    On Error GoTo EH
    For Each cell In Target.Cells
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

' ===========================================================================
' SetCellBoldSafe
' ===========================================================================
' Назначение:
'   Безопасно управляет жирным начертанием одной ячейки.
'   Это локальная версия `SetRangeBoldSafe`, удобная для точечной разметки шапок и заголовков.
' Основные действия:
'   - Проверяет, входит ли ячейка в объединённую область.
'   - Применяет изменение к корректной целевой ячейке.
'   - При ошибке возвращает диагностическое сообщение с адресом.
' Параметры:
'   Target - ячейка, стиль которой нужно изменить.
'   isBold - признак, нужно ли сделать текст жирным.
'   tag - служебная метка этапа форматирования.
' Эффект выполнения:
'   Процедура меняет формат шрифта одной ячейки или её объединённой области.
Private Sub SetCellBoldSafe(ByVal Target As Range, ByVal isBold As Boolean, Optional ByVal tag As String = "")
    On Error GoTo EH
    If Target.MergeCells Then
        Target.MergeArea.Cells(1, 1).Font.Bold = isBold
    Else
        Target.Font.Bold = isBold
    End If
    Exit Sub

EH:
    Err.Raise Err.Number, "SetCellBoldSafe", "Bold stage '" & tag & "' at " & Target.Address(False, False) & ": " & Err.Description
End Sub

' ===========================================================================
' BuildValidationDateFormula
' ===========================================================================
' Назначение:
'   Преобразует дату в строку, пригодную для передачи в правило валидации Excel.
'   Excel ожидает серийный номер даты, а не форматированную строку.
' Основные действия:
'   - Берёт только целую серийную часть даты.
'   - Возвращает её в текстовом виде для дальнейшей подстановки в Validation.
' Параметры:
'   d - дата, которую нужно подготовить для правила проверки.
' Возвращаемое значение:
'   Строковое представление серийного номера даты Excel.
Private Function BuildValidationDateFormula(ByVal d As Date) As String
    BuildValidationDateFormula = CStr(CLng(d))
End Function

' ===========================================================================
' BuildValidationDecimalFormula
' ===========================================================================
' Назначение:
'   Готовит числовую границу для текстовой формулы валидации.
'   Учитывает локальный десятичный разделитель Excel, чтобы формулы работали
'   независимо от региональных настроек системы.
' Основные действия:
'   - Получает локальный десятичный разделитель приложения Excel.
'   - Нормализует текст числа под этот разделитель.
' Параметры:
'   valueNum - число, которое нужно вставить в формулу проверки.
' Возвращаемое значение:
'   Строка с числом в локально совместимом виде.
Private Function BuildValidationDecimalFormula(ByVal valueNum As Double) As String
    Dim decSep As String
    Dim txt As String

    decSep = Application.International(xlDecimalSeparator)
    txt = Trim$(CStr(valueNum))
    txt = Replace$(txt, ".", decSep)
    txt = Replace$(txt, ",", decSep)
    BuildValidationDecimalFormula = txt
End Function

' ===========================================================================
' ApplyDateValidation
' ===========================================================================
' Назначение:
'   Назначает правилу ввода ячейки ограничение по диапазону дат.
'   Используется там, где дата может редактироваться пользователем вручную.
' Основные действия:
'   - Удаляет предыдущее правило Validation.
'   - Создаёт новое правило типа `xlValidateDate`.
'   - Ограничивает допустимый диапазон между minDate и maxDate.
' Параметры:
'   Target - ячейка или диапазон, для которого назначается правило.
'   minDate - минимально допустимая дата.
'   maxDate - максимально допустимая дата.
' Эффект выполнения:
'   Excel начинает блокировать ввод дат за пределами допустимого интервала.
Private Sub ApplyDateValidation(ByVal Target As Range, ByVal minDate As Date, ByVal maxDate As Date)
    With Target.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:=CLng(minDate), Formula2:=CLng(maxDate)
        .IgnoreBlank = True
    End With
End Sub

' ===========================================================================
' ApplyTimeValidation
' ===========================================================================
' Назначение:
'   Ограничивает ввод диапазоном корректного времени суток.
'   Используется в ячейках, где пользователь может вручную редактировать время.
' Основные действия:
'   - Удаляет текущее правило Validation.
'   - Создаёт правило `xlValidateTime` для диапазона от 00:00:00 до 23:59:59.
' Параметры:
'   Target - ячейка или диапазон, куда разрешён ввод времени.
' Эффект выполнения:
'   Пользователь не сможет ввести значение вне стандартного суточного времени.
Private Sub ApplyTimeValidation(ByVal Target As Range)
    With Target.Validation
        .Delete
        .Add Type:=xlValidateTime, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:=0, Formula2:=86399# / 86400#
        .IgnoreBlank = True
    End With
End Sub

' ===========================================================================
' ApplyDecimalValidation
' ===========================================================================
' Назначение:
'   Назначает стандартную числовую валидацию для десятичного значения.
'   Подходит для ячеек, где Excel хранит именно число, а не текстовую форму числа.
' Основные действия:
'   - Удаляет прежнее правило проверки.
'   - Назначает правило `xlValidateDecimal` с диапазоном minVal..maxVal.
' Параметры:
'   Target - ячейка или диапазон, для которого включается проверка.
'   minVal - минимально допустимое значение.
'   maxVal - максимально допустимое значение.
' Эффект выполнения:
'   Excel ограничивает ввод десятичных чисел заданным интервалом.
Private Sub ApplyDecimalValidation(ByVal Target As Range, ByVal minVal As Double, ByVal maxVal As Double)
    With Target.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:=minVal, Formula2:=maxVal
        .IgnoreBlank = True
    End With
End Sub

' ===========================================================================
' ApplyDecimalTextValidation
' ===========================================================================
' Назначение:
'   Назначает валидацию для ячейки, где десятичное число вводится как текст.
'   Такой подход нужен, когда важно сохранить пользовательский контроль над форматом
'   ввода, но при этом не допустить мусорных значений.
' Основные действия:
'   - Строит адрес ячейки для формулы Validation.
'   - Готовит локально совместимые текстовые границы диапазона.
'   - Создаёт пользовательскую формулу, которая допускает пустую ячейку
'     либо корректное десятичное число в заданных пределах.
' Параметры:
'   Target - ячейка или диапазон с текстовым представлением числа.
'   minVal - нижняя допустимая граница.
'   maxVal - верхняя допустимая граница.
' Эффект выполнения:
'   Excel проверяет текстовый ввод как десятичное число без жёсткого перевода ячейки в numeric-only режим.
Private Sub ApplyDecimalTextValidation(ByVal Target As Range, ByVal minVal As Double, ByVal maxVal As Double)
    Dim addr As String
    Dim q As String
    Dim minTxt As String
    Dim maxTxt As String
    Dim valueExpr As String

    addr = Target.Cells(1, 1).Address(False, False)
    q = Chr$(34)
    minTxt = BuildValidationDecimalFormula(minVal)
    maxTxt = BuildValidationDecimalFormula(maxVal)
    valueExpr = "VALUE(SUBSTITUTE(" & addr & "," & q & "." & q & "," & q & "," & q & "))"

    With Target.Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, _
            Formula1:="=OR(LEN(TRIM(" & addr & "))=0,AND(ISNUMBER(IFERROR(" & valueExpr & ",FALSE)),IFERROR(" & valueExpr & ",0)>=" & minTxt & ",IFERROR(" & valueExpr & ",0)<=" & maxTxt & "))"
        .IgnoreBlank = True
    End With
End Sub

' ===========================================================================
' ApplyWholeValidation
' ===========================================================================
' Назначение:
'   Назначает стандартную валидацию для целых чисел.
'   Используется в ячейках, где Excel должен принимать только целочисленные значения.
' Основные действия:
'   - Удаляет прежнюю проверку.
'   - Назначает правило `xlValidateWholeNumber` с заданными границами.
' Параметры:
'   Target - ячейка или диапазон проверки.
'   minVal - минимально допустимое целое значение.
'   maxVal - максимально допустимое целое значение.
' Эффект выполнения:
'   Excel запрещает ввод дробных и выходящих за диапазон значений.
Private Sub ApplyWholeValidation(ByVal Target As Range, ByVal minVal As Long, ByVal maxVal As Long)
    With Target.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:=minVal, Formula2:=maxVal
        .IgnoreBlank = True
    End With
End Sub

' ===========================================================================
' ApplyWholeTextValidation
' ===========================================================================
' Назначение:
'   Назначает валидацию для текстового поля, которое должно содержать целое число.
'   Полезно там, где значение хранится как текст, но по смыслу должно оставаться
'   допустимым числовым идентификатором без дробной части.
' Основные действия:
'   - Строит формульное выражение для проверки содержимого ячейки.
'   - Разрешает пустое значение либо целое число в заданных границах.
' Параметры:
'   Target - ячейка или диапазон с текстовым числовым вводом.
'   minVal - нижняя граница допустимого целого значения.
'   maxVal - верхняя граница допустимого целого значения.
' Эффект выполнения:
'   Excel отбрасывает нечисловой текст и дробные значения в целочисленных полях.
Private Sub ApplyWholeTextValidation(ByVal Target As Range, ByVal minVal As Long, ByVal maxVal As Long)
    Dim addr As String
    Dim valueExpr As String

    addr = Target.Cells(1, 1).Address(False, False)
    valueExpr = "VALUE(" & addr & ")"

    With Target.Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, _
            Formula1:="=OR(LEN(TRIM(" & addr & "))=0,AND(ISNUMBER(IFERROR(" & valueExpr & ",FALSE)),IFERROR(" & valueExpr & ",0)=INT(IFERROR(" & valueExpr & ",0)),IFERROR(" & valueExpr & ",0)>=" & CStr(minVal) & ",IFERROR(" & valueExpr & ",0)<=" & CStr(maxVal) & "))"
        .IgnoreBlank = True
    End With
End Sub

' ===========================================================================
' RefreshWorkbookColors
' ===========================================================================
' Назначение:
'   Полностью пересобирает цветовое оформление всей книги.
'   Это главный сценарий синхронизации интерфейса после изменения цветовых настроек,
'   очистки листов или других операций, которые могли нарушить консистентность оформления.
' Основные действия:
'   - Получает ссылки на все рабочие листы книги.
'   - Снимает защиту и временно отключает события и перерисовку.
'   - Гарантирует наличие блока настройки цветов и считывает актуальную палитру.
'   - Вызывает специализированные процедуры перекраски для каждого листа.
'   - Возвращает защиту листов и стандартные настройки Excel.
' Эффект выполнения:
'   Процедура массово обновляет внешний вид книги, не меняя бизнес-данные пользователя.
Public Sub RefreshWorkbookColors()
    On Error GoTo EH

    Dim wsDisclaimer As Worksheet, wsIn As Worksheet, wsOut As Worksheet, wsHist As Worksheet, wsMRS As Worksheet, wsNotes As Worksheet
    Set wsDisclaimer = ThisWorkbook.Worksheets(1)
    Set wsIn = ThisWorkbook.Worksheets(2)
    Set wsOut = ThisWorkbook.Worksheets(3)
    Set wsHist = ThisWorkbook.Worksheets(4)
    Set wsMRS = ThisWorkbook.Worksheets(5)
    Set wsNotes = ThisWorkbook.Worksheets(6)

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
    wsNotes.Unprotect UW(49, 49, 52, 55, 48, 57)

    EnsureColorSettings wsIn
    ReadAllColors wsIn, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrMrsOrderUnconf, clrHeader

    RefreshDisclaimerSheetColors wsDisclaimer, clrLocked
    RefreshInputSheetColors wsIn, wsHist, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrMrsOrderUnconf, clrHeader
    RefreshResultSheetColors wsOut, clrLocked
    RefreshHistorySheetColors wsHist, clrLocked, clrEditHasColor, clrEditable, clrMrsOrder, clrMrsOrderUnconf, clrHeader
    RefreshMRSSheetColors wsMRS, clrLocked, clrEditHasColor, clrEditable, clrMrsHeader, clrMrsSub, clrMrsOrder, clrMrsOrderUnconf, clrHeader
    RefreshNotesSheetColors wsNotes, clrLocked, clrEditHasColor, clrEditable, clrHeader

Cleanup:
    ' Возвращаем защиту листов через тихие helper-процедуры,
    ' чтобы cleanup не создавал вложенный обработчик ошибок.
    ProtectSheetQuiet wsDisclaimer
    ProtectSheetQuiet wsIn
    ProtectSheetQuiet wsOut
    ProtectSheetQuiet wsHist
    ProtectSheetQuiet wsMRS
    ProtectSheetQuiet wsNotes
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
EH:
    Resume Cleanup
End Sub

' ===========================================================================
' PickCellColor
' ===========================================================================
' Назначение:
'   Открывает встроенный диалог Excel для выбора цвета одной строки цветовых настроек.
'   Это общий внутренний helper для всех кнопок выбора цвета на листе ввода.
' Основные действия:
'   - Находит лист "Ввод" и временно снимает его защиту.
'   - Выделяет ячейку нужной строки в колонке образцов цвета.
'   - Открывает системный диалог выбора узора/цвета.
'   - Если цвет изменился, запускает перекраску всей книги.
' Параметры:
'   rowNum - номер строки в блоке цветовых настроек.
' Эффект выполнения:
'   Пользователь может изменить один из цветов интерфейса, после чего книга синхронизируется.
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

' ===========================================================================
' PickColorLocked
' ===========================================================================
' Назначение:
'   Точка входа для выбора цвета заблокированных ячеек.
' Эффект выполнения:
'   Запускает общий сценарий выбора цвета для строки заблокированного состояния.
Public Sub PickColorLocked()
    PickCellColor CLR_ROW_LOCKED
End Sub

' ===========================================================================
' PickColorEditable
' ===========================================================================
' Назначение:
'   Точка входа для выбора цвета редактируемых ячеек.
' Эффект выполнения:
'   Запускает общий сценарий выбора цвета для редактируемого состояния.
Public Sub PickColorEditable()
    PickCellColor CLR_ROW_EDITABLE
End Sub

' ===========================================================================
' PickColorMrsHeader
' ===========================================================================
' Назначение:
'   Точка входа для выбора цвета заголовков дат на листе MRS.
' Эффект выполнения:
'   Запускает общий сценарий выбора соответствующего цвета.
Public Sub PickColorMrsHeader()
    PickCellColor CLR_ROW_MRS_HEADER
End Sub

' ===========================================================================
' PickColorMrsSub
' ===========================================================================
' Назначение:
'   Точка входа для выбора цвета строк бригад на листе MRS.
' Эффект выполнения:
'   Запускает общий сценарий выбора соответствующего цвета.
Public Sub PickColorMrsSub()
    PickCellColor CLR_ROW_MRS_SUBHEADER
End Sub

' ===========================================================================
' PickColorMrsOrder
' ===========================================================================
' Назначение:
'   Точка входа для выбора цвета подтверждённых заказов.
' Эффект выполнения:
'   Запускает общий сценарий выбора соответствующего цвета.
Public Sub PickColorMrsOrder()
    PickCellColor CLR_ROW_MRS_ORDER
End Sub

' ===========================================================================
' PickColorMrsOrderUnconf
' ===========================================================================
' Назначение:
'   Точка входа для выбора цвета неподтверждённых заказов.
' Эффект выполнения:
'   Запускает общий сценарий выбора соответствующего цвета.
Public Sub PickColorMrsOrderUnconf()
    PickCellColor CLR_ROW_MRS_ORDER_UNCONF
End Sub

' ===========================================================================
' PickColorHeader
' ===========================================================================
' Назначение:
'   Точка входа для выбора цвета шапок и служебных заголовков.
' Эффект выполнения:
'   Запускает общий сценарий выбора соответствующего цвета.
Public Sub PickColorHeader()
    PickCellColor CLR_ROW_HEADER
End Sub

' ===========================================================================
' UW
' ===========================================================================
' Назначение:
'   Собирает Unicode-строку из массива числовых кодов символов.
'   Эта функция используется для хранения русскоязычных строк в виде кодов, чтобы
'   модуль был устойчивее к проблемам кодировки в VBA-источнике.
' Параметры:
'   codes - набор числовых кодов символов.
' Возвращаемое значение:
'   Готовая строка Unicode, собранная из переданных кодов.
Public Function UW(ParamArray codes() As Variant) As String
    Dim i As Long
    For i = LBound(codes) To UBound(codes)
        UW = UW & ChrW(CLng(codes(i)))
    Next i
End Function

' ===========================================================================
' ZQ
' ===========================================================================
' Назначение:
'   Выполняет внутреннюю проверку допустимости запуска основных сценариев.
'   Функция сравнивает скрытый маркер на листе ввода с вычисленным эталоном и
'   используется как ранний предохранитель перед крупными операциями.
' Основные действия:
'   - Читает служебное значение с листа ввода.
'   - Собирает ожидаемую строку из закодированного массива.
'   - Сравнивает фактическое и ожидаемое значения побайтно.
' Возвращаемое значение:
'   True, если внутренняя проверка пройдена; иначе False.
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

' ===========================================================================
' NumForFormula
' ===========================================================================
' Назначение:
'   Преобразует число в текст для безопасной подстановки в Excel-формулу.
'   Независимо от локали Excel итоговая строка использует точку как разделитель,
'   что удобно при ручной сборке формульных выражений.
' Параметры:
'   val - числовое значение для вставки в формулу.
' Возвращаемое значение:
'   Строка с числом в формульно-совместимом виде.
Private Function NumForFormula(ByVal val As Double) As String
    NumForFormula = Replace$(Trim$(Str$(val)), Application.DecimalSeparator, ".")
End Function

' ===========================================================================
' GenerateAndAppendHistory
' ===========================================================================
' Назначение:
'   Выполняет основной расчёт калькулятора рабочего времени.
'   Это центральная процедура модуля: она читает параметры с листа "Ввод",
'   строит цепочку операций на листе результата, а затем переносит итог в историю.
' Основные действия:
'   - Проверяет защитный маркер запуска через ZQ().
'   - Снимает защиту с листов, подготавливает цветовые настройки и очищает прошлый результат.
'   - Читает даты, время, паузы, параметры обеда, список работников и строки операций.
'   - Для каждой операции и каждого выбранного исполнителя рассчитывает старт, окончание,
'     длительность и связанные служебные поля.
'   - Формирует текстовый блок Z7 и переносит итоговый набор строк в историю.
'   - Обновляет стартовые дату и время на листе ввода по фактическому окончанию цепочки.
'   - Возвращает защиту листов и сообщает пользователю количество сформированных строк.
' Эффект выполнения:
'   На листе результата строится новая таблица расчёта, история дополняется новыми строками,
'   а лист ввода синхронизируется с финальной точкой времени последней операции.
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

    ' На время расчёта снимаем защиту и отключаем события,
    ' чтобы служебные записи не запускали обработчики повторно.
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

    ' Результат пересобирается с нуля, поэтому старые строки сначала удаляются.
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
    lunchDurMin = val(Replace(CStr(wsIn.Range("B12").Value), ",", "."))
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

    ' Основной цикл проходит по строкам операций на листе ввода.
    ' Каждая найденная операция затем разворачивается в одну или несколько строк результата.
    Do While opRow <= 23
        Dim opName As String
        opName = Trim$(CStr(wsIn.Cells(opRow, 9).Value))
        If opName = "" Then Exit Do

        opNum = opNum + 1

        Dim durVal As Double, durUnit As String, timeMode As String
        durVal = val(Replace(CStr(wsIn.Cells(opRow, 11).Value), ",", "."))
        If durVal < 0 Then durVal = 0
        durUnit = NormalizeUnit(wsIn.Cells(opRow, 12).Value)
        timeMode = LCase$(Trim$(CStr(wsIn.Cells(opRow, 13).Value)))

        Dim breakVal As Double, breakUnit As String
        breakVal = val(Replace(CStr(wsIn.Cells(opRow, 14).Value), ",", "."))
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
        ' Внутри операции создаём отдельную строку для каждого выбранного исполнителя.
        ' Первый исполнитель задаёт базовые значения, остальные наследуют часть полей формулами.
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

                ' Заполняем строку результата как самостоятельную запись операции,
                ' но повторяющиеся значения внутри одной группы связываем через формулы.
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

    ' После построения данных приводим лист результата к финальному виду:
    ' форматы времени, границы и ширины нужны уже до переноса блока в историю.
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
    ' Блок Z7 формирует текстовое резюме расчёта для выдачи и для последующей записи в историю.
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

    ' Сначала переносим готовый расчёт в историю, затем обновляем оформление
    ' и сохраняем стартовую точку для следующего запуска.
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
    ' Quiet-helper'ы здесь нужны для безопасного завершения даже в аварийном сценарии:
    ' сначала возвращается защита листов, затем восстанавливается редактируемая ячейка паузы.
    ProtectSheetQuiet wsOut
    ProtectSheetQuiet wsHist
    Application.EnableEvents = True
    RestorePauseInputCellQuiet wsIn, clrEditHasColor, clrEditable
    ProtectSheetQuiet wsIn
    If Len(errMsg) > 0 Then
        MsgBox "Error: " & errMsg, vbCritical
    End If
    Exit Sub

EH:
    errMsg = Err.Description
    Resume Cleanup
End Sub

' ===========================================================================
' EnsureSheetHeaders
' ===========================================================================
' Назначение:
'   Восстанавливает служебные заголовки на листе результата и базовый заголовок истории.
'   Процедура нужна перед генерацией новых расчётных строк, чтобы структура листов
'   всегда была предсказуемой и не зависела от предыдущих операций очистки.
' Основные действия:
'   - Заполняет строку заголовков результата по всем рабочим колонкам A:V.
'   - Выделяет шапку жирным шрифтом.
'   - Восстанавливает заголовок "История" на листе истории.
' Параметры:
'   wsOut - лист результата.
'   wsHist - лист истории.
' Эффект выполнения:
'   Процедура меняет только шапки листов и их форматирование.
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

' ===========================================================================
' ClearResultAndHistory
' ===========================================================================
' Назначение:
'   Сбрасывает расчётную часть книги к исходному рабочему состоянию.
'   Используется, когда нужно полностью очистить результат и историю перед новым расчётом.
' Основные действия:
'   - Проверяет допустимость запуска через ZQ().
'   - Сбрасывает базовые дату и время на листе ввода.
'   - Очищает лист результата.
'   - Очищает лист истории.
'   - После очистки запускает полную перекраску книги.
' Эффект выполнения:
'   Пользователь получает пустые листы результата и истории при сохранении структуры книги.
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
    ' После полного сброса защита листов восстанавливается единообразно
    ' тем же quiet-helper'ом, что и в других cleanup-блоках модуля.
    ProtectSheetQuiet wsIn
    ProtectSheetQuiet wsR
    ProtectSheetQuiet wsH
    Exit Sub
EH:
    Resume Cleanup
End Sub

' ===========================================================================
' ClearResultArea
' ===========================================================================
' Назначение:
'   Удаляет все расчётные строки с листа результата, оставляя шапку.
'   Процедура используется как в полном сбросе книги, так и перед построением нового результата.
' Основные действия:
'   - Определяет фактическую нижнюю границу использованной области.
'   - Временно отключает события Excel.
'   - Удаляет строки результата, начиная со второй.
' Параметры:
'   wsOut - лист результата.
' Эффект выполнения:
'   Лист результата очищается от старых расчётных данных.
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

' ===========================================================================
' ClearHistoryArea
' ===========================================================================
' Назначение:
'   Удаляет все блоки истории, оставляя служебную верхнюю часть листа.
'   Используется при полном сбросе истории перед новым циклом работы.
' Основные действия:
'   - Определяет фактический конец данных на листе истории.
'   - Временно отключает события.
'   - Удаляет все строки, начиная с четвёртой.
' Параметры:
'   wsHist - лист истории.
'   clrLocked - базовый цвет листа; в текущей реализации параметр передаётся для согласованности интерфейса вызова.
' Эффект выполнения:
'   История очищается от сохранённых расчётных блоков.
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

' ===========================================================================
' SyncWorkerIdInputs
' ===========================================================================
' Назначение:
'   Подготавливает блок ввода работников под указанное количество исполнителей.
'   Процедура управляет видимостью строк, подписями и стилями в области ввода идентификаторов.
' Основные действия:
'   - Нормализует количество работников в допустимый диапазон 1..10.
'   - Обеспечивает наличие и актуальность цветовых настроек.
'   - Восстанавливает заголовок блока исполнителей.
'   - Проходит по 10 возможным строкам и показывает только нужное количество.
'   - Для каждой строки вызывает отдельную процедуру оформления.
' Параметры:
'   wsIn - лист "Ввод".
'   workerCount - количество работников, которое должно быть активно на листе.
' Эффект выполнения:
'   Лист ввода адаптируется под текущее число исполнителей без изменения бизнес-логики расчёта.
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

' ===========================================================================
' SyncOperationRows
' ===========================================================================
' Назначение:
'   Приводит блок ввода операций к нужному количеству активных строк.
'   Это одна из ключевых процедур динамического интерфейса листа "Ввод".
' Основные действия:
'   - Нормализует допустимое число операций в диапазон 1..20.
'   - Считывает текущие цветовые настройки и временно отключает события Excel.
'   - Показывает нужное число строк операций и скрывает лишние.
'   - Подставляет стандартные значения в пустые ячейки операций.
'   - Синхронизирует связанные колонки с типами длительностей и перерывов.
' Параметры:
'   wsIn - лист "Ввод".
'   opCount - число строк операций, которое должно остаться активным.
' Эффект выполнения:
'   Рабочая область операций на листе ввода становится согласованной с текущим числом операций.
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

' ===========================================================================
' SanitizeWorkerIdCell
' ===========================================================================
' Назначение:
'   Очищает ввод идентификатора исполнителя непосредственно в ячейке Excel.
'   Процедура приводит пользовательский ввод к внутреннему формату идентификатора
'   и оставляет в ячейке либо корректное значение, либо пустоту.
' Основные действия:
'   - Нормализует строку через `NormalizeWorkerIdText`.
'   - Устанавливает формат отображения идентификатора как восьмизначного числа.
'   - Очищает ячейку, если после нормализации данных не осталось.
' Параметры:
'   targetCell - ячейка с идентификатором исполнителя.
' Эффект выполнения:
'   Пользовательский ввод в поле исполнителя приводится к единому безопасному формату.
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

' ===========================================================================
' StyleWorkerInputRow
' ===========================================================================
' Назначение:
'   Оформляет одну строку блока исполнителей на листе "Ввод".
'   Процедура используется при синхронизации количества работников и отвечает
'   одновременно за видимость, блокировку и стиль метки с полем ввода.
' Основные действия:
'   - Оформляет колонку подписи исполнителя как служебную и заблокированную.
'   - Для поля идентификатора назначает либо редактируемый стиль, либо скрывает/очищает строку.
'   - Настраивает границы, формат номера и доступность редактирования.
' Параметры:
'   wsIn - лист "Ввод".
'   rowNum - номер строки блока исполнителей.
'   isVisible - должна ли строка быть активной и видимой для пользователя.
'   clrLocked - цвет заблокированного состояния.
'   clrEditHasColor - признак пользовательского цвета редактируемых ячеек.
'   clrEditable - цвет редактируемого состояния.
' Эффект выполнения:
'   Конкретная строка исполнителя приводится к нужному состоянию интерфейса.
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

' ===========================================================================
' GetWorkerValue
' ===========================================================================
' Назначение:
'   Возвращает нормализованное значение исполнителя для дальнейшей записи в расчёт.
'   Если пользователь не ввёл идентификатор, функция может подставить номер работника
'   как резервное значение.
' Основные действия:
'   - Читает текст идентификатора из блока исполнителей.
'   - Нормализует его с учётом fallback на номер работника.
'   - Возвращает либо числовой идентификатор, либо индекс исполнителя.
' Параметры:
'   wsIn - лист "Ввод".
'   workerIndex - порядковый номер исполнителя в блоке.
' Возвращаемое значение:
'   Нормализованное значение исполнителя для записи в результат и историю.
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

' ===========================================================================
' GetWorkerNumberFormat
' ===========================================================================
' Назначение:
'   Определяет формат отображения идентификатора исполнителя.
'   Нужна для сохранения различия между пустым значением и осмысленным восьмизначным кодом.
' Основные действия:
'   - Читает и нормализует идентификатор исполнителя.
'   - Возвращает либо общий формат, либо строгий формат `00000000`.
' Параметры:
'   wsIn - лист "Ввод".
'   workerIndex - номер строки исполнителя.
' Возвращаемое значение:
'   Строка формата Excel для отображения идентификатора исполнителя.
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

' ===========================================================================
' NormalizeWorkerIdText
' ===========================================================================
' Назначение:
'   Приводит ввод идентификатора исполнителя к каноническому виду.
'   Функция оставляет только цифры, ограничивает длину восемью символами и при
'   необходимости дополняет значение ведущими нулями.
' Основные действия:
'   - Удаляет все нецифровые символы.
'   - При пустом вводе либо возвращает пустую строку, либо использует fallback.
'   - Ограничивает длину идентификатора и выравнивает его до восьми знаков.
' Параметры:
'   rawText - исходный текстовый ввод.
'   workerIndex - номер исполнителя, используемый как резервное значение.
'   fallbackToIndex - признак, нужно ли подставлять номер исполнителя при пустом вводе.
' Возвращаемое значение:
'   Строка идентификатора в стандартизированном восьмизначном виде либо пустая строка.
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

' ===========================================================================
' NormalizeDecimal
' ===========================================================================
' Назначение:
'   Нормализует произвольный текстовый ввод десятичного числа.
'   Функция удаляет лишние символы, сохраняет только первую десятичную часть и
'   приводит результат к компактной форме с запятой.
' Основные действия:
'   - Проходит по символам исходной строки.
'   - Оставляет цифры и только один десятичный разделитель.
'   - Ограничивает целую часть шестью знаками, а дробную двумя.
' Параметры:
'   rawText - исходный текст, введённый пользователем.
' Возвращаемое значение:
'   Строка с нормализованным десятичным числом либо пустая строка.
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

' ===========================================================================
' NormalizeRizInput
' ===========================================================================
' Назначение:
'   Подготавливает ввод для поля РИЗ.
'   Значение сначала нормализуется как десятичное число, после чего к нему
'   добавляется единица измерения в человекочитаемом виде.
' Основные действия:
'   - Использует `NormalizeDecimal` как базовый механизм очистки.
'   - Если число валидно, дописывает текст единицы измерения.
' Параметры:
'   rawText - исходный ввод пользователя.
' Возвращаемое значение:
'   Нормализованный текст для поля РИЗ либо пустая строка.
Public Function NormalizeRizInput(ByVal rawText As String) As String
    Dim normalized As String
    normalized = NormalizeDecimal(rawText)
    If normalized = "" Then
        NormalizeRizInput = ""
    Else
        NormalizeRizInput = normalized & UW(32, 1052, 1086, 1084)
    End If
End Function

' ===========================================================================
' NormalizeKInput
' ===========================================================================
' Назначение:
'   Подготавливает ввод для коэффициента K.
'   В отличие от РИЗ здесь не добавляется текстовая приставка: сохраняется только
'   нормализованное числовое представление.
' Параметры:
'   rawText - исходный текст пользователя.
' Возвращаемое значение:
'   Нормализованное строковое представление коэффициента K.
Public Function NormalizeKInput(ByVal rawText As String) As String
    NormalizeKInput = NormalizeDecimal(rawText)
End Function

' ===========================================================================
' SanitizeDecimalEntryCell
' ===========================================================================
' Назначение:
'   Применяет нормализацию десятичного ввода непосредственно к ячейке Excel.
'   Используется в обработчиках изменений истории, MRS и входных полей.
' Основные действия:
'   - Временно переводит ячейку в текстовый формат.
'   - Нормализует значение через `NormalizeDecimal`.
'   - Либо записывает корректное число, либо очищает ячейку, либо подставляет ноль.
' Параметры:
'   targetCell - редактируемая ячейка с десятичным значением.
'   blankAsZero - нужно ли заменять пустой результат на ноль.
' Эффект выполнения:
'   Ячейка получает согласованное числовое значение без лишних символов.
Public Sub SanitizeDecimalEntryCell(ByVal targetCell As Range, Optional ByVal blankAsZero As Boolean = True)
    Dim cleaned As String

    If targetCell Is Nothing Then Exit Sub

    targetCell.NumberFormat = "@"
    cleaned = NormalizeDecimal(CStr(targetCell.Value))

    If cleaned = "" Then
        If blankAsZero Then
            targetCell.Value = 0
        Else
            targetCell.ClearContents
        End If
    Else
        targetCell.Value = val(Replace(cleaned, ",", "."))
    End If
End Sub

' ===========================================================================
' SanitizeWholeTextCell
' ===========================================================================
' Назначение:
'   Очищает ячейку, которая должна содержать целое число в текстовом виде.
'   Это удобно для полей, где важен не арифметический тип Excel, а набор допустимых цифр.
' Основные действия:
'   - Извлекает из исходного ввода только цифры.
'   - Ограничивает длину результата по `maxDigits`.
'   - Записывает очищенный текст обратно или заменяет пустоту на ноль.
' Параметры:
'   targetCell - ячейка с целочисленным текстовым вводом.
'   maxDigits - максимальное количество цифр, которое допускается сохранить.
'   blankAsZero - нужно ли заменять пустой результат на строку "0".
' Эффект выполнения:
'   Ячейка очищается от посторонних символов и приводится к допустимому виду.
Public Sub SanitizeWholeTextCell(ByVal targetCell As Range, ByVal maxDigits As Long, Optional ByVal blankAsZero As Boolean = True)
    Dim digits As String

    If targetCell Is Nothing Then Exit Sub

    digits = DigitsOnly(CStr(targetCell.Value))
    If maxDigits > 0 And Len(digits) > maxDigits Then
        digits = Left$(digits, maxDigits)
    End If

    targetCell.NumberFormat = "@"
    If digits = "" Then
        If blankAsZero Then
            targetCell.Value = "0"
        Else
            targetCell.ClearContents
        End If
    Else
        targetCell.Value = digits
    End If
End Sub

' ===========================================================================
' SanitizeTimeCell
' ===========================================================================
' Назначение:
'   Нормализует ввод времени в пользовательской ячейке.
'   Процедура принимает как текстовые представления времени, так и внутренний
'   числовой формат Excel, после чего приводит значение к единому виду.
' Основные действия:
'   - Пробует интерпретировать содержимое как текст времени или число Excel.
'   - Устанавливает формат отображения `hh:mm:ss`.
'   - При невалидном вводе очищает ячейку.
' Параметры:
'   targetCell - ячейка со временем.
' Эффект выполнения:
'   В ячейке остаётся корректное время Excel либо пустое значение.
Public Sub SanitizeTimeCell(ByVal targetCell As Range)
    Dim rawText As String
    Dim parsedTime As Date

    If targetCell Is Nothing Then Exit Sub

    rawText = Trim$(CStr(targetCell.Value))
    targetCell.NumberFormat = "hh:mm:ss"

    If rawText = "" Then
        targetCell.ClearContents
        Exit Sub
    End If

    On Error GoTo InvalidValue
    If IsNumeric(targetCell.Value) Then
        parsedTime = CDbl(targetCell.Value) - Fix(CDbl(targetCell.Value))
    Else
        parsedTime = TimeValue(rawText)
    End If
    On Error GoTo 0

    targetCell.Value = parsedTime
    Exit Sub

Cleanup:
    targetCell.ClearContents
    Exit Sub

InvalidValue:
    Resume Cleanup
End Sub

' ===========================================================================
' HandleHistorySheetChange
' ===========================================================================
' Назначение:
'   Централизованно обрабатывает пользовательские изменения на листе "История".
'   После переноса логики из модуля листа эта процедура вызывается из
'   обработчика `Workbook_SheetChange` и позволяет оставить сам лист без VBA-кода.
' Основные действия:
'   - Временно отключает события и снимает защиту с листа.
'   - При изменении статуса в колонке B перекрашивает всю строку заказа.
'   - Для редактируемых десятичных колонок запускает санитизацию ввода.
'   - Для редактируемых целочисленных колонок оставляет только допустимые цифры.
'   - Возвращает защиту листа и повторно включает события Excel.
' Параметры:
'   ws - лист истории, в котором произошло изменение.
'   Target - диапазон изменённых пользователем ячеек.
' Эффект выполнения:
'   Исправляет пользовательский ввод и поддерживает корректное цветовое состояние строк истории.
Public Sub HandleHistorySheetChange(ByVal ws As Worksheet, ByVal Target As Range)
    On Error GoTo SafeExit
    If ws Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub

    Application.EnableEvents = False
    ws.Unprotect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)

    Dim bRange As Range
    Set bRange = Intersect(Target, ws.Columns(2))
    If Not bRange Is Nothing Then
        Dim clrMrsOrder As Long, clrMrsOrderUnconf As Long
        GetOrderColors clrMrsOrder, clrMrsOrderUnconf
        Dim bCell As Range
        For Each bCell In bRange.Cells
            If bCell.Row >= 4 Then
                ColorOrderRow ws, bCell.Row, 22, clrMrsOrder, clrMrsOrderUnconf
            End If
        Next bCell
    End If

    Dim decimalRange As Range
    Set decimalRange = Intersect(Target, Union(ws.Columns(12), ws.Columns(5)))
    If Not decimalRange Is Nothing Then
        Dim cell As Range
        For Each cell In decimalRange.Cells
            If cell.Row < 4 Then GoTo NextHistoryDecimalCell
            If cell.Locked Then GoTo NextHistoryDecimalCell
            If cell.HasFormula Then GoTo NextHistoryDecimalCell
            SanitizeDecimalEntryCell cell
NextHistoryDecimalCell:
        Next cell
    End If

    Dim wholeRange As Range
    Set wholeRange = Intersect(Target, Union(ws.Columns(7), ws.Columns(16)))
    If Not wholeRange Is Nothing Then
        Dim wholeCell As Range
        For Each wholeCell In wholeRange.Cells
            If wholeCell.Row < 4 Then GoTo NextHistoryWholeCell
            If wholeCell.Locked Then GoTo NextHistoryWholeCell
            If wholeCell.HasFormula Then GoTo NextHistoryWholeCell
            SanitizeWholeTextCell wholeCell, 8
NextHistoryWholeCell:
        Next wholeCell
    End If

SafeExit:
    ' Защита листа возвращается отдельным helper'ом,
    ' чтобы безопасный выход не переключал внешний error handler.
    ProtectSheetQuiet ws
    Application.EnableEvents = True
End Sub

' ===========================================================================
' HandleMRSSheetChange
' ===========================================================================
' Назначение:
'   Централизованно обрабатывает изменения на листе "Парсинг MRS".
'   Процедура выполняет ту же роль, что и обработчик истории, но учитывает
'   структуру и набор редактируемых колонок именно для листа MRS.
' Основные действия:
'   - Временно снимает защиту и отключает события.
'   - При изменении статуса в колонке B перекрашивает строку заказа.
'   - Нормализует десятичные поля MRS.
'   - Нормализует и проверяет временную колонку.
'   - Возвращает защиту листа и включает события обратно.
' Параметры:
'   ws - лист MRS, в котором произошло изменение.
'   Target - диапазон изменённых ячеек.
' Эффект выполнения:
'   Поддерживает чистый ввод данных и актуальное цветовое оформление строк на листе MRS.
Public Sub HandleMRSSheetChange(ByVal ws As Worksheet, ByVal Target As Range)
    On Error GoTo SafeExit
    If ws Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub

    Application.EnableEvents = False
    ws.Unprotect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)

    Dim bRange As Range
    Set bRange = Intersect(Target, ws.Columns(2))
    If Not bRange Is Nothing Then
        Dim clrMrsOrder As Long, clrMrsOrderUnconf As Long
        GetOrderColors clrMrsOrder, clrMrsOrderUnconf
        Dim bCell As Range
        For Each bCell In bRange.Cells
            If bCell.Row >= 4 Then
                ColorOrderRow ws, bCell.Row, 14, clrMrsOrder, clrMrsOrderUnconf
            End If
        Next bCell
    End If

    Dim decimalRange As Range
    Set decimalRange = Intersect(Target, Union(ws.Columns(9), ws.Columns(13)))
    If Not decimalRange Is Nothing Then
        Dim decimalCell As Range
        For Each decimalCell In decimalRange.Cells
            If decimalCell.Row < 4 Then GoTo NextMrsDecimalCell
            If decimalCell.Locked Then GoTo NextMrsDecimalCell
            If decimalCell.HasFormula Then GoTo NextMrsDecimalCell
            SanitizeDecimalEntryCell decimalCell
NextMrsDecimalCell:
        Next decimalCell
    End If

    Dim timeRange As Range
    Set timeRange = Intersect(Target, ws.Columns(7))
    If Not timeRange Is Nothing Then
        Dim timeCell As Range
        For Each timeCell In timeRange.Cells
            If timeCell.Row < 4 Then GoTo NextMrsTimeCell
            If timeCell.Locked Then GoTo NextMrsTimeCell
            If timeCell.HasFormula Then GoTo NextMrsTimeCell
            SanitizeTimeCell timeCell
NextMrsTimeCell:
        Next timeCell
    End If

SafeExit:
    ' Тот же quiet-cleanup, что и у истории:
    ' на выходе лист просто возвращается в защищённое состояние.
    ProtectSheetQuiet ws
    Application.EnableEvents = True
End Sub

' ===========================================================================
' ProtectSheetQuiet
' ===========================================================================
' Назначение:
'   Тихо возвращает защиту одного листа в cleanup-блоках.
'   Процедура нужна там, где повторный сбой при `Protect` не должен ломать
'   основной сценарий завершения и подменять уже сохранённую ошибку.
' Основные действия:
'   - Проверяет, передан ли объект листа.
'   - Пытается включить для него стандартный режим защиты.
'   - Локально гасит возможную ошибку только внутри helper-процедуры.
' Параметры:
'   ws - лист, для которого нужно восстановить защиту.
' Эффект выполнения:
'   Лист по возможности возвращается в защищённое состояние без влияния на внешний обработчик ошибок.
Private Sub ProtectSheetQuiet(ByVal ws As Worksheet)
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Protect Chr(49) & Chr(49) & Chr(52) & Chr(55) & Chr(48) & Chr(57)
    End If
    On Error GoTo 0
End Sub

' ===========================================================================
' ProtectWorkbookQuiet
' ===========================================================================
' Назначение:
'   Тихо возвращает структурную защиту книги.
'   Используется в сценариях импорта MRS, где защита книги временно снимается
'   и затем должна быть восстановлена без вложенных обработчиков ошибок.
' Основные действия:
'   - Проверяет наличие объекта книги.
'   - Вызывает `Workbook.Protect` с переданными флагами структуры и окон.
'   - Изолирует возможную ошибку внутри процедуры.
' Параметры:
'   wb - книга, для которой нужно включить защиту.
'   protectStructure - признак защиты структуры книги.
'   protectWindows - признак защиты окон книги.
' Эффект выполнения:
'   Книга по возможности возвращается в защищённое состояние без срыва cleanup-логики.
Private Sub ProtectWorkbookQuiet(ByVal wb As Workbook, ByVal protectStructure As Boolean, ByVal protectWindows As Boolean)
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Protect UW(49, 49, 52, 55, 48, 57), protectStructure, protectWindows
    End If
    On Error GoTo 0
End Sub

' ===========================================================================
' CloseWorkbookQuiet
' ===========================================================================
' Назначение:
'   Безопасно закрывает временно открытую книгу-источник.
'   Helper нужен в cleanup-блоках импорта, где книга могла быть не открыта
'   или уже закрыта к моменту завершения процедуры.
' Основные действия:
'   - Проверяет наличие объекта книги.
'   - Пытается закрыть её без сохранения.
'   - Не даёт вторичной ошибке прервать основную ветку cleanup.
' Параметры:
'   wb - временная книга, которую нужно закрыть.
' Эффект выполнения:
'   Временный workbook по возможности освобождается без вмешательства во внешний сценарий ошибок.
Private Sub CloseWorkbookQuiet(ByVal wb As Workbook)
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close False
    End If
    On Error GoTo 0
End Sub

' ===========================================================================
' MergeRangeQuiet
' ===========================================================================
' Назначение:
'   Выполняет объединение диапазона без переназначения внешнего error handler.
'   Нужен для служебных областей, где `Merge` может вызвать исключение, если
'   диапазон уже приведён к нужному состоянию предыдущим запуском.
' Основные действия:
'   - Проверяет наличие диапазона.
'   - Пытается выполнить `Merge`.
'   - Завершает локальный тихий блок обработки ошибок.
' Параметры:
'   targetRange - диапазон, который требуется объединить.
' Эффект выполнения:
'   Диапазон по возможности объединяется, не меняя логику внешнего cleanup/exception flow.
Private Sub MergeRangeQuiet(ByVal targetRange As Range)
    On Error Resume Next
    If Not targetRange Is Nothing Then
        targetRange.Merge
    End If
    On Error GoTo 0
End Sub

' ===========================================================================
' RestorePauseInputCellQuiet
' ===========================================================================
' Назначение:
'   Тихо восстанавливает редактируемое состояние ячейки паузы на листе ввода.
'   Используется в финальной части расчёта, где cleanup обязан вернуть ячейку
'   в рабочее состояние даже после аварийного выхода из основной процедуры.
' Основные действия:
'   - Проверяет наличие листа ввода.
'   - Снимает блокировку с ячейки `N4`.
'   - Применяет к ней текущий стиль редактируемого поля.
' Параметры:
'   wsIn - лист ввода.
'   clrEditHasColor - признак пользовательского цвета редактируемых ячеек.
'   clrEditable - цвет редактируемой ячейки.
' Эффект выполнения:
'   Поле паузы возвращается в доступное состояние без риска сорвать cleanup новой ошибкой.
Private Sub RestorePauseInputCellQuiet(ByVal wsIn As Worksheet, ByVal clrEditHasColor As Boolean, ByVal clrEditable As Long)
    On Error Resume Next
    If Not wsIn Is Nothing Then
        wsIn.Cells(4, 14).Locked = False
        ApplyEditableStyle wsIn.Cells(4, 14), clrEditHasColor, clrEditable
    End If
    On Error GoTo 0
End Sub

' ===========================================================================
' DigitsOnly
' ===========================================================================
' Назначение:
'   Извлекает из строки только цифры.
'   Это один из самых базовых helper-ов модуля, используемый в нормализации
'   идентификаторов, целочисленных полей и списков участников.
' Основные действия:
'   - Последовательно проходит по всем символам строки.
'   - Добавляет в результат только символы от `0` до `9`.
' Параметры:
'   rawText - исходный текст, который нужно очистить.
' Возвращаемое значение:
'   Строка, содержащая только цифры из исходного значения.
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

' ===========================================================================
' AppendParticipantsRun
' ===========================================================================
' Назначение:
'   Добавляет в итоговую строку один диапазон выбранных участников.
'   Используется при обратной сборке нормализованного списка работников из булевого массива.
' Основные действия:
'   - Формирует токен либо одиночного номера, либо диапазона `start-end`.
'   - Добавляет его в результирующую строку через запятую.
' Параметры:
'   result - строка-накопитель результата.
'   startNum - первый номер диапазона.
'   endNum - последний номер диапазона.
' Эффект выполнения:
'   В строку спецификации добавляется ещё один компактный блок выбора участников.
Private Sub AppendParticipantsRun(ByRef result As String, ByVal startNum As Long, ByVal endNum As Long)
    Dim token As String

    If startNum = endNum Then
        token = CStr(startNum)
    Else
        token = CStr(startNum) & "-" & CStr(endNum)
    End If

    If result <> "" Then result = result & ","
    result = result & token
End Sub

' ===========================================================================
' NormalizeParticipantsSpec
' ===========================================================================
' Назначение:
'   Приводит список участников операции к каноническому строковому формату.
'   Поддерживает одиночные номера и диапазоны, например `1,2,4-6`, и удаляет
'   некорректные варианты ввода.
' Основные действия:
'   - Очищает строку от пробелов и альтернативных разделителей.
'   - Проверяет синтаксис чисел и диапазонов.
'   - Строит внутренний булев массив выбранных работников.
'   - Собирает результат обратно в компактную нормализованную запись.
' Параметры:
'   rawText - исходный пользовательский список участников.
'   maxWorkerCount - максимальное допустимое число работников.
' Возвращаемое значение:
'   Нормализованная строка выбора участников либо пустая строка при ошибке или пустом списке.
Public Function NormalizeParticipantsSpec(ByVal rawText As String, ByVal maxWorkerCount As Long) As String
    Dim spec As String
    Dim i As Long
    Dim ch As String
    Dim token As Variant
    Dim dashCount As Long
    Dim bounds() As String
    Dim fromN As Long, toN As Long, tmp As Long
    Dim n As Long
    Dim hasAny As Boolean
    Dim inRun As Boolean
    Dim runStart As Long, runEnd As Long
    Dim selected() As Boolean

    If maxWorkerCount < 1 Then maxWorkerCount = 1
    If maxWorkerCount > 10 Then maxWorkerCount = 10

    spec = Trim$(rawText)
    spec = Replace$(spec, vbTab, "")
    spec = Replace$(spec, " ", "")
    spec = Replace$(spec, ";", ",")
    spec = Replace$(spec, ".", ",")
    spec = Replace$(spec, ChrW(&H2013), "-")
    spec = Replace$(spec, ChrW(&H2014), "-")
    spec = Replace$(spec, ChrW(&H2212), "-")

    If spec = "" Then
        NormalizeParticipantsSpec = ""
        Exit Function
    End If

    For i = 1 To Len(spec)
        ch = Mid$(spec, i, 1)
        If Not ((ch >= "0" And ch <= "9") Or ch = "," Or ch = "-") Then
            GoTo InvalidSpec
        End If
    Next i

    ReDim selected(1 To maxWorkerCount)
    For Each token In Split(spec, ",")
        token = Trim$(CStr(token))
        If token = "" Then GoTo ContinueToken

        dashCount = Len(token) - Len(Replace$(token, "-", ""))
        If dashCount > 1 Then
            GoTo InvalidSpec
        ElseIf dashCount = 1 Then
            bounds = Split(CStr(token), "-")
            If UBound(bounds) <> 1 Then
                GoTo InvalidSpec
            End If
            If bounds(0) = "" Or bounds(1) = "" Then
                GoTo InvalidSpec
            End If
            If DigitsOnly(bounds(0)) <> bounds(0) Or DigitsOnly(bounds(1)) <> bounds(1) Then
                GoTo InvalidSpec
            End If

            fromN = CLng(bounds(0))
            toN = CLng(bounds(1))
            If fromN < 1 Or fromN > maxWorkerCount Or toN < 1 Or toN > maxWorkerCount Then
                GoTo InvalidSpec
            End If
            If fromN > toN Then
                tmp = fromN
                fromN = toN
                toN = tmp
            End If

            For n = fromN To toN
                selected(n) = True
                hasAny = True
            Next n
        Else
            If DigitsOnly(CStr(token)) <> CStr(token) Then
                GoTo InvalidSpec
            End If

            n = CLng(token)
            If n < 1 Or n > maxWorkerCount Then
                GoTo InvalidSpec
            End If
            selected(n) = True
            hasAny = True
        End If
ContinueToken:
    Next token

    If Not hasAny Then
        NormalizeParticipantsSpec = ""
        Exit Function
    End If

    For n = 1 To maxWorkerCount
        If selected(n) Then
            If Not inRun Then
                runStart = n
                runEnd = n
                inRun = True
            ElseIf n = runEnd + 1 Then
                runEnd = n
            Else
                AppendParticipantsRun NormalizeParticipantsSpec, runStart, runEnd
                runStart = n
                runEnd = n
            End If
        ElseIf inRun Then
            AppendParticipantsRun NormalizeParticipantsSpec, runStart, runEnd
            inRun = False
        End If
    Next n

    If inRun Then
        AppendParticipantsRun NormalizeParticipantsSpec, runStart, runEnd
    End If
    Exit Function

InvalidSpec:
    NormalizeParticipantsSpec = ""
End Function

' ===========================================================================
' SanitizeParticipantsCell
' ===========================================================================
' Назначение:
'   Применяет нормализацию списка участников прямо к ячейке Excel.
'   Используется в строках операций на листе ввода.
' Основные действия:
'   - Читает список участников из ячейки.
'   - Прогоняет его через `NormalizeParticipantsSpec`.
'   - Записывает нормализованную строку обратно или очищает ячейку.
' Параметры:
'   targetCell - ячейка со спецификацией участников.
'   maxWorkerCount - максимальное количество допустимых работников.
' Эффект выполнения:
'   Ячейка получает компактный и синтаксически корректный список участников.
Public Sub SanitizeParticipantsCell(ByVal targetCell As Range, ByVal maxWorkerCount As Long)
    Dim normalized As String

    If targetCell Is Nothing Then Exit Sub

    normalized = NormalizeParticipantsSpec(CStr(targetCell.Value), maxWorkerCount)
    targetCell.NumberFormat = "@"
    If normalized = "" Then
        targetCell.ClearContents
    Else
        targetCell.Value = normalized
    End If
End Sub

' ===========================================================================
' TimeValueDefault
' ===========================================================================
' Назначение:
'   Безопасно извлекает временную часть из произвольного значения.
'   Если значение нельзя интерпретировать как время, функция возвращает запасной вариант.
' Основные действия:
'   - Проверяет числовой формат Excel.
'   - Пытается распознать дату/время или текстовое время.
'   - При ошибке возвращает `Fallback`.
' Параметры:
'   raw - исходное значение из ячейки или параметра.
'   Fallback - запасное время по умолчанию.
' Возвращаемое значение:
'   Корректное значение времени Excel.
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

' ===========================================================================
' ConvertDurationToDays
' ===========================================================================
' Назначение:
'   Переводит длительность из минут или часов в долю суток Excel.
'   Внутри расчётов модуля это базовый формат для сложения дат и времени.
' Параметры:
'   valueNum - длительность в исходных единицах.
'   unitName - нормализованное имя единицы измерения.
' Возвращаемое значение:
'   Длительность в формате доли суток Excel.
Private Function ConvertDurationToDays(ByVal valueNum As Double, ByVal unitName As String) As Double
    If unitName = "hour" Then
        ConvertDurationToDays = valueNum / 24#
    Else
        ConvertDurationToDays = valueNum / 1440#
    End If
End Function

' ===========================================================================
' NormalizeUnit
' ===========================================================================
' Назначение:
'   Приводит текст единицы измерения к внутреннему стандарту модуля.
'   Это нужно, чтобы расчёты не зависели от разных вариантов написания часов и минут.
' Основные действия:
'   - Переводит текст в нижний регистр.
'   - Сводит варианты записи часа к значению `hour`.
'   - Все остальные случаи трактует как минуты.
' Параметры:
'   rawUnit - исходный текст единицы измерения.
' Возвращаемое значение:
'   Каноническое имя единицы измерения: `hour` или `min`.
Private Function NormalizeUnit(ByVal rawUnit As Variant) As String
    Dim unitName As String
    unitName = LCase$(Trim$(CStr(rawUnit)))
    If unitName = "hour" Or unitName = LCase$(UW(1095, 1072, 1089)) Or unitName = LCase$(UW(1095)) Then
        NormalizeUnit = "hour"
    Else
        NormalizeUnit = "min"
    End If
End Function

' ===========================================================================
' ParseLunchParams
' ===========================================================================
' Назначение:
'   Подготавливает параметры обеденного перерыва к расчётам времени.
'   На выходе процедура даёт уже нормализованные значения времени начала обеда,
'   наличие второго интервала и длительность в долях суток.
' Основные действия:
'   - Преобразует сырые значения времени в Excel Time.
'   - Ограничивает длительность разумными рамками.
'   - Вычисляет числовое значение длительности в долях суток.
' Параметры:
'   lunch1Raw - исходное значение первого обеденного интервала.
'   lunch2Raw - исходное значение второго обеденного интервала.
'   lunchDurMin - длительность обеда в минутах; при выходе может быть скорректирована.
'   lunch1 - сюда возвращается нормализованное время первого интервала.
'   lunch2 - сюда возвращается нормализованное время второго интервала.
'   hasLunch2 - флаг наличия второго интервала.
'   lunchDurDays - сюда возвращается длительность в долях суток.
' Эффект выполнения:
'   Процедура заполняет подготовленные значения для всех последующих расчётных helper-функций.
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

' ===========================================================================
' ComputeEndWithLunch
' ===========================================================================
' Назначение:
'   Вычисляет время окончания операции с учётом возможного пересечения обеда.
'   Если операция пересекает обеденный интервал, конец сдвигается на длительность
'   перерыва, а вызывающий код получает отметку о факте такого пересечения.
' Основные действия:
'   - Вычисляет базовое окончание без учёта обеда.
'   - Проверяет пересечение первого обеденного интервала.
'   - При наличии второго интервала отдельно проверяет и его.
'   - При пересечении добавляет длительность обеда к окончанию.
' Параметры:
'   opStart - время начала операции.
'   durationDays - длительность операции в долях суток.
'   lunch1 - начало первого обеденного интервала.
'   lunch2 - начало второго обеденного интервала.
'   hasLunch2 - флаг наличия второго интервала.
'   lunchDurDays - длительность обеда в долях суток.
'   crossedLunch - выходной флаг, показывающий факт пересечения обеда.
' Возвращаемое значение:
'   Время окончания операции с учётом сдвига из-за обеда.
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

' ===========================================================================
' ShiftStartOutOfLunch
' ===========================================================================
' Назначение:
'   Сдвигает старт операции, если он попал внутрь обеденного интервала.
'   Нужна для того, чтобы расчёт не начинал новую операцию посреди перерыва.
' Основные действия:
'   - Проверяет попадание времени начала в первый обеденный интервал.
'   - При наличии второго интервала делает аналогичную проверку и для него.
'   - Если старт попал внутрь обеда, возвращает конец соответствующего интервала.
' Параметры:
'   startDateTime - исходное расчётное время начала.
'   lunch1 - начало первого обеденного интервала.
'   lunch2 - начало второго обеденного интервала.
'   hasLunch2 - флаг наличия второго интервала.
'   lunchDurDays - длительность обеда в долях суток.
' Возвращаемое значение:
'   Скорректированное время старта, не попадающее внутрь обеда.
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

' ===========================================================================
' BuildLunchShiftFormulaStr
' ===========================================================================
' Назначение:
'   Собирает Excel-формулу для сдвига старта операции за пределы обеда.
'   Формула нужна в истории, где часть расчётов должна жить не в VBA, а прямо в ячейках Excel.
' Основные действия:
'   - Строит условие для первого обеденного интервала.
'   - При наличии второго интервала добавляет и второе условие.
'   - Возвращает выражение, которое либо оставляет исходное время, либо переносит его
'     на конец соответствующего интервала.
' Параметры:
'   rawTimeExpr - выражение Excel для исходного времени.
'   lunch1 - начало первого обеда.
'   lunch2 - начало второго обеда.
'   lunchDurMin - длительность обеда в минутах.
' Возвращаемое значение:
'   Текст Excel-формулы, корректирующей стартовое время.
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

' ===========================================================================
' BuildEndFormulasStr
' ===========================================================================
' Назначение:
'   Генерирует Excel-формулы для даты и времени окончания операции.
'   Процедура нужна для листа истории, где вычисление конца операции должно
'   оставаться живым и пересчитываемым прямо в ячейках.
' Основные действия:
'   - Собирает общую математическую часть формулы из старта, длительности и поправок на обед.
'   - Добавляет логику первого обеденного интервала.
'   - При наличии второго интервала расширяет формулу и для него.
'   - Возвращает две отдельные формулы: для даты и для времени.
' Параметры:
'   startDateExpr - выражение Excel для даты начала.
'   startTimeExpr - выражение Excel для времени начала.
'   durExpr - выражение Excel для длительности.
'   unitDiv - делитель перевода длительности в долю суток.
'   lunch1 - начало первого обеденного интервала.
'   lunch2 - начало второго обеденного интервала.
'   lunchDurMin - длительность обеда в минутах.
'   outEndDateFormula - сюда возвращается формула даты окончания.
'   outEndTimeFormula - сюда возвращается формула времени окончания.
' Эффект выполнения:
'   Процедура подготавливает две строковые формулы для записи в ячейки истории.
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

' ===========================================================================
' BuildLunchIconFormulaStr
' ===========================================================================
' Назначение:
'   Собирает Excel-формулу для индикатора пересечения обеда.
'   Эта формула позже записывается в ячейки истории и возвращает визуальную метку,
'   если операция была сдвинута обедом или пересекла обеденный интервал.
' Основные действия:
'   - Формирует текстовые выражения для первого интервала обеда.
'   - При наличии второго интервала добавляет в формулу и его.
'   - Возвращает готовую строку `IF(...)` для вставки в ячейку Excel.
' Параметры:
'   startTimeExpr - выражение Excel для времени начала.
'   durExpr - выражение Excel для длительности.
'   unitDiv - делитель, переводящий длительность в долю суток.
'   lunch1 - начало первого обеда.
'   lunch2 - начало второго обеда.
'   lunchDurMin - длительность обеда в минутах.
'   daStr - текстовая метка, возвращаемая формулой при совпадении условия.
' Возвращаемое значение:
'   Готовая строка Excel-формулы для индикатора обеда.
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

' ===========================================================================
' TimePart
' ===========================================================================
' Назначение:
'   Возвращает только временную часть значения Date.
'   Это небольшой helper для построения расчётов по времени внутри суток.
' Параметры:
'   dateTimeValue - дата и время Excel.
' Возвращаемое значение:
'   Дробная часть числа даты, соответствующая времени.
Private Function TimePart(ByVal dateTimeValue As Date) As Double
    TimePart = CDbl(dateTimeValue) - Fix(CDbl(dateTimeValue))
End Function

' ===========================================================================
' IsWorkerSelected
' ===========================================================================
' Назначение:
'   Проверяет, входит ли конкретный работник в строковую спецификацию участников.
'   Если спецификация пуста, функция трактует это как выбор всех работников.
' Основные действия:
'   - Разбирает список по запятым.
'   - Поддерживает одиночные номера и диапазоны вида `3-5`.
'   - Возвращает истину при первом совпадении.
' Параметры:
'   spec - спецификация участников операции.
'   workerIndex - номер проверяемого работника.
' Возвращаемое значение:
'   True, если работник входит в выбор; иначе False.
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

' ===========================================================================
' JoinOperationNames
' ===========================================================================
' Назначение:
'   Собирает уникальные названия операций из листа результата в одну строку.
'   Используется для текстового блока Z7, где нужен компактный перечень работ.
' Основные действия:
'   - Проходит по строкам результата в указанном диапазоне.
'   - Сохраняет только уникальные имена операций.
'   - Объединяет их через запятую.
' Параметры:
'   wsOut - лист результата.
'   firstRow - первая строка диапазона операций.
'   lastRow - последняя строка диапазона операций.
' Возвращаемое значение:
'   Строка с перечнем уникальных операций.
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

' ===========================================================================
' AppendResultToHistory
' ===========================================================================
' Назначение:
'   Переносит свежесформированный расчёт с листа результата на лист истории.
'   Это один из наиболее важных сценариев модуля: именно здесь временный расчёт
'   превращается в постоянную запись истории с редактируемыми колонками, формулами,
'   цветами статуса и дополнительным блоком Z7.
' Основные действия:
'   - Находит позицию, куда будет добавлен новый блок истории.
'   - Создаёт заголовок заказа со статусом и данными о времени формирования.
'   - Копирует шапку и строки расчёта с листа результата.
'   - Перестраивает формулы дат, времени начала и окончания уже в контексте истории.
'   - Назначает форматы, границы, редактируемые поля и правила валидации.
'   - Добавляет следом текстовый блок Z7.
'   - Перекрашивает добавленный блок по статусу заказа.
' Параметры:
'   wsHist - лист истории, куда добавляется новый блок.
'   wsOut - лист результата, содержащий свежий расчёт.
'   wsIn - лист ввода, содержащий номер и имя заказа.
'   lastDataRow - последняя строка расчётных данных на листе результата.
'   lastZ7Row - последняя строка текстового блока Z7 на листе результата.
'   workerCount - количество работников в текущем расчёте.
'   lunch1 - начало первого обеденного интервала.
'   lunch2 - начало второго обеденного интервала.
'   lunchDurMin - длительность обеда в минутах.
'   primaryIsMin - флаг основной единицы длительности расчёта.
'   clrEditHasColor - признак пользовательского цвета редактируемых ячеек.
'   clrEditable - цвет редактируемых ячеек.
'   clrMrsOrder - цвет подтверждённого заказа.
'   clrMrsOrderUnconf - цвет неподтверждённого заказа.
' Эффект выполнения:
'   На листе истории появляется новый полноформатный блок заказа, готовый к просмотру, правкам и экспорту.
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
    ' Новый блок истории всегда вставляется после предыдущего,
    ' а между блоками сохраняется пустая разделительная строка.
    NextRow = wsHist.Cells(wsHist.Rows.Count, 2).End(xlUp).Row + 1
    If NextRow < 4 Then NextRow = 4

    If NextRow > 4 Then NextRow = NextRow + 1

    Dim orderNum As String, orderName As String
    orderNum = Trim$(CStr(wsIn.Range("B3").Value))
    orderName = Trim$(CStr(wsIn.Range("B4").Value))
    
    ' Верхняя строка блока хранит статус заказа, номер и краткое описание,
    ' чтобы в истории можно было быстро отличать один расчёт от другого.
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

    ' Далее копируем шапку результата и все расчётные строки,
    ' а уже потом поверх них перестраиваем формулы под контекст листа истории.
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

    ' Ищем конец предыдущего блока истории, чтобы новый блок мог
    ' корректно продолжить общую временную цепочку заказов.
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

    ' Старт и конец каждой строки истории задаются уже формулами,
    ' чтобы при ручной корректировке пауз и длительностей блок пересчитывался автоматически.
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
    wsHist.Range(wsHist.Cells(headerRow, 1), wsHist.Cells(headerRow, 22)).WrapText = True
    wsHist.Rows(headerRow).AutoFit

    wsHist.Range(wsHist.Cells(headerRow, 7), wsHist.Cells(dataEndRow, 7)).Borders(7).LineStyle = xlContinuous
    wsHist.Range(wsHist.Cells(headerRow, 7), wsHist.Cells(dataEndRow, 7)).Borders(7).Weight = 4
    wsHist.Range(wsHist.Cells(headerRow, 21), wsHist.Cells(dataEndRow, 21)).Borders(10).LineStyle = xlContinuous
    wsHist.Range(wsHist.Cells(headerRow, 21), wsHist.Cells(dataEndRow, 21)).Borders(10).Weight = 4

    NextRow = NextRow + (lastDataRow - 1) + 1

    Dim z7Start As Long
    ' Текстовое резюме Z7 переносится в историю отдельным подблоком,
    ' чтобы его можно было читать рядом с соответствующим расчётом.
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
    ' После копирования возвращаем редактируемость только тем колонкам,
    ' которые допускают ручную корректировку в истории.
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
                        wsHist.Cells(rr, editCol(ec)).NumberFormat = "@"
                        ApplyDecimalTextValidation wsHist.Cells(rr, editCol(ec)), 0, 999999.99
                    Case 7, 16
                        Dim wholeText As String
                        wholeText = DigitsOnly(CStr(wsHist.Cells(rr, editCol(ec)).Text))
                        wsHist.Cells(rr, editCol(ec)).NumberFormat = "@"
                        If wholeText <> "" Then
                            wsHist.Cells(rr, editCol(ec)).Value = wholeText
                        End If
                        ApplyWholeTextValidation wsHist.Cells(rr, editCol(ec)), 0, 99999999
                End Select
            End If
SkipCell:
        Next rr
    Next ec

    wsHist.Cells(dataStartRow, 15).Locked = False
    If clrEditHasColor Then
        wsHist.Cells(dataStartRow, 15).Interior.Color = clrEditable
        wsHist.Cells(dataStartRow, 15).Font.Color = GetContrastColor(clrEditable)
    Else
        wsHist.Cells(dataStartRow, 15).Interior.Pattern = xlNone
        wsHist.Cells(dataStartRow, 15).Font.Color = 0
    End If
    ApplyDateValidation wsHist.Cells(dataStartRow, 15), DateSerial(2000, 1, 1), DateSerial(2099, 12, 31)

    wsHist.Range(wsHist.Cells(dataStartRow, 5), wsHist.Cells(dataEndRow, 5)).NumberFormat = "@"
    wsHist.Range(wsHist.Cells(dataStartRow, 7), wsHist.Cells(dataEndRow, 7)).NumberFormat = "@"
    wsHist.Range(wsHist.Cells(dataStartRow, 12), wsHist.Cells(dataEndRow, 12)).NumberFormat = "@"
    wsHist.Range(wsHist.Cells(dataStartRow, 16), wsHist.Cells(dataEndRow, 16)).NumberFormat = "@"

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

    wsHist.Columns("A:A").ColumnWidth = 1
    wsHist.Columns("B:B").ColumnWidth = 3
    wsHist.Columns("C:C").ColumnWidth = 63
    wsHist.Columns("D:D").ColumnWidth = 8
    wsHist.Columns("E:E").ColumnWidth = 10
    wsHist.Columns("F:F").ColumnWidth = 10
    wsHist.Columns("G:G").ColumnWidth = 12
    wsHist.Columns("H:H").ColumnWidth = 1
    wsHist.Columns("I:I").ColumnWidth = 1
    wsHist.Columns("J:J").ColumnWidth = 1
    wsHist.Columns("K:K").ColumnWidth = 1
    wsHist.Columns("L:L").ColumnWidth = 10
    wsHist.Columns("M:M").ColumnWidth = 1
    wsHist.Columns("N:N").ColumnWidth = 1
    wsHist.Columns("O:O").ColumnWidth = 15
    wsHist.Columns("P:P").ColumnWidth = 16
    wsHist.Columns("Q:Q").ColumnWidth = 1
    wsHist.Columns("R:R").ColumnWidth = 14
    wsHist.Columns("S:S").ColumnWidth = 14
    wsHist.Columns("T:T").ColumnWidth = 14
    wsHist.Columns("U:U").ColumnWidth = 14
    wsHist.Columns("V:V").ColumnWidth = 8
End Sub

' ===========================================================================
' IsNumericOrder
' ===========================================================================
' Назначение:
'   Проверяет, можно ли считать строку корректным числовым номером заказа.
'   Используется при фильтрации входных данных MRS, чтобы отбросить строки,
'   которые не похожи на реальные номера заказов.
' Параметры:
'   s - строка-кандидат на номер заказа.
' Возвращаемое значение:
'   True, если строка состоит только из цифр; иначе False.
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

' ===========================================================================
' LoadMRS
' ===========================================================================
' Назначение:
'   Загружает внешний Excel-файл с данными MRS и раскладывает его в рабочий лист книги.
'   Это один из самых крупных сценариев модуля: он отвечает и за чтение источника,
'   и за фильтрацию записей, и за их группировку по датам, заказам и бригадам.
' Основные действия:
'   - Очищает текущий лист MRS и подготавливает цветовые настройки книги.
'   - Запрашивает у пользователя исходный файл и открывает его в режиме ReadOnly.
'   - Считывает исходный диапазон в массив для ускоренной обработки.
'   - Отфильтровывает валидные строки заказов и раскладывает их по служебным массивам.
'   - Собирает набор дат, при необходимости показывает диалог выбора дат.
'   - Группирует строки по бригадам и заказам, при необходимости показывает диалог выбора бригад.
'   - Записывает результат на лист MRS, оформляет заголовки, строки данных и цвета.
' Параметры:
'   Процедура не принимает аргументов и работает с активной книгой как с контекстом.
' Эффект выполнения:
'   Лист "Парсинг MRS" полностью перестраивается на основе выбранного внешнего файла.
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
    ' Цвета и параметры обеда нужны и для оформления MRS,
    ' и для расчётных формул, которые будут строиться в импортируемых блоках.
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
    lunchDurMin = val(Replace(CStr(wsIn.Range("B12").Value), ",", "."))
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
    ' Исходный файл сначала раскладывается в массивы.
    ' Это заметно быстрее, чем читать и преобразовывать каждую ячейку напрямую на листе.
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
    ' Отдельно собираем список дат из исходника,
    ' чтобы пользователь мог ограничить импорт только нужными днями.
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
    ' При наличии нескольких дат показываем множественный выбор,
    ' чтобы не тащить в лист MRS лишние блоки.
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
                If lbDate.selected(dIdx) Then
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
    ' Каждая выбранная дата формирует свой верхнеуровневый блок:
    ' дата, затем бригады внутри даты, затем отдельные заказы внутри бригад.
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
        ' Сначала строим связи "заказ -> состав работников" и запоминаем
        ' самое раннее время старта заказа, чтобы дальше правильно группировать и сортировать блоки.
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
    ' Заказы с одинаковым набором исполнителей сворачиваются в одну бригаду.
    ' Ключ бригады строится как отсортированный список табельных номеров.
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
    ' Если в выбранной дате несколько бригад,
    ' пользователь может импортировать только нужные составы.
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
                If lb.selected(bDisp) Then
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

        ' Внутри бригады заказы выводятся по фактическому времени старта,
        ' а между ними при необходимости вставляется редактируемая пауза.
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
                wsMRS.Cells(outRow, 9).NumberFormat = "@"
                wsMRS.Cells(outRow, 9).Value = 0
                wsMRS.Cells(outRow, 9).Locked = False
                If clrEditHasColor Then
                    wsMRS.Cells(outRow, 9).Interior.Color = clrEditable
                Else
                    wsMRS.Cells(outRow, 9).Interior.Pattern = xlNone
                End If
                ApplyDecimalTextValidation wsMRS.Cells(outRow, 9), 0, 999999.99
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
                wsMRS.Cells(outRow, 13).NumberFormat = "@"
                wsMRS.Cells(outRow, 13).Value = Round(initMin, 2)
                wsMRS.Cells(outRow, 13).Locked = False
                If clrEditHasColor Then
                    wsMRS.Cells(outRow, 13).Interior.Color = clrEditable
                Else
                    wsMRS.Cells(outRow, 13).Interior.Pattern = xlNone
                End If
                ApplyDecimalTextValidation wsMRS.Cells(outRow, 13), 0, 999999.99

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

    wsMRS.Columns("A:A").ColumnWidth = 1
    wsMRS.Columns("B:B").ColumnWidth = 4
    wsMRS.Columns("C:C").ColumnWidth = 13
    wsMRS.Columns("D:D").ColumnWidth = 45
    wsMRS.Columns("E:E").ColumnWidth = 12
    wsMRS.Columns("F:F").ColumnWidth = 15
    wsMRS.Columns("G:G").ColumnWidth = 17
    wsMRS.Columns("H:H").ColumnWidth = 15
    wsMRS.Columns("I:I").ColumnWidth = 17
    wsMRS.Columns("J:J").ColumnWidth = 19
    wsMRS.Columns("K:K").ColumnWidth = 19
    wsMRS.Columns("L:L").ColumnWidth = 18
    wsMRS.Columns("M:M").ColumnWidth = 18
    wsMRS.Columns("N:N").ColumnWidth = 10

    MsgBox UW(1055, 1072, 1088, 1089, 1080, 1085, 1075, 32, 1079, 1072, 1074, 1077, 1088, 1096, 1077, 1085, 32, 1091, 1089, 1087, 1077, 1096, 1085, 1086, 46), vbInformation

    Dim errMsg As String
Done:
    ' Очистка импорта вынесена в единые quiet-helper'ы:
    ' отдельно закрывается временный источник, затем возвращается защита книги и листа MRS.
    CloseWorkbookQuiet srcWb
    ProtectWorkbookQuiet ThisWorkbook, True, False
    ProtectSheetQuiet wsMRS
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = calcState
    Application.StatusBar = False
    If Len(errMsg) > 0 Then
        MsgBox UW(1054, 1096, 1080, 1073, 1082, 1072, 58, 32) & errMsg, vbCritical
    End If
    Exit Sub
EH:
    errMsg = stage & " | " & Err.Description
    Resume Done
End Sub

' ===========================================================================
' QuickSortLong
' ===========================================================================
' Назначение:
'   Быстро сортирует массив целых чисел по возрастанию.
'   Используется при обработке дат и других служебных числовых наборов.
' Параметры:
'   arr - сортируемый массив Long.
'   first - левая граница сортировки.
'   last - правая граница сортировки.
' Эффект выполнения:
'   Массив сортируется на месте алгоритмом QuickSort.
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

' ===========================================================================
' QuickSortVariantStrings
' ===========================================================================
' Назначение:
'   Быстро сортирует строковые значения, хранящиеся в Variant-массиве.
'   Нужна там, где ключи словарей извлекаются как Variant и должны быть
'   упорядочены перед дальнейшей группировкой.
' Параметры:
'   arr - массив строковых значений в Variant-форме.
'   first - левая граница сортировки.
'   last - правая граница сортировки.
' Эффект выполнения:
'   Массив сортируется на месте алгоритмом QuickSort.
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

' ===========================================================================
' QuickSortDoubleString
' ===========================================================================
' Назначение:
'   Сортирует числовой массив и синхронно переставляет связанный строковый массив.
'   Применяется там, где время и строковый ключ должны сохранять взаимное соответствие.
' Параметры:
'   arrD - массив числовых ключей сортировки.
'   arrS - связанный строковый массив.
'   first - левая граница сортировки.
'   last - правая граница сортировки.
' Эффект выполнения:
'   Оба массива переставляются синхронно, сохраняя связь между элементами.
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

' ===========================================================================
' QuickSortIndices
' ===========================================================================
' Назначение:
'   Сортирует массив индексов по составному набору признаков связанных массивов.
'   Такой подход удобен, когда нужно менять только порядок ссылок на данные,
'   не переставляя сами массивы с датами, временем и строками.
' Параметры:
'   idx - массив индексов, который сортируется.
'   arrSDate - массив дат старта.
'   arrSTime - массив времени старта.
'   arrOp - массив названий или ключей операций.
'   arrWID - массив идентификаторов работников.
'   first - левая граница сортировки.
'   last - правая граница сортировки.
' Эффект выполнения:
'   Индексный массив упорядочивается согласно логике `CompareIndices`.
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

' ===========================================================================
' CompareIndices
' ===========================================================================
' Назначение:
'   Сравнивает одну запись с опорным набором сортировочных значений.
'   Используется внутри `QuickSortIndices` для составной сортировки по времени,
'   операции и идентификатору работника.
' Параметры:
'   idxA - индекс сравниваемой записи.
'   pA - опорное числовое значение даты и времени.
'   pB - опорный строковый ключ операции.
'   pC - опорный строковый ключ работника.
'   arrSDate - массив дат старта.
'   arrSTime - массив времени старта.
'   arrOp - массив операций.
'   arrWID - массив идентификаторов работников.
' Возвращаемое значение:
'   `-1`, `0` или `1` в зависимости от результата сравнения.
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

' ===========================================================================
' ClearMRS
' ===========================================================================
' Назначение:
'   Полностью очищает лист MRS и возвращает книгу к согласованному состоянию.
'   Используется перед новым импортом и как отдельная пользовательская команда очистки.
' Основные действия:
'   - Проверяет допустимость запуска через внутренний предохранитель.
'   - Считывает текущие цвета интерфейса.
'   - Удаляет рабочие строки листа MRS.
'   - Запускает общую перекраску книги после очистки.
' Эффект выполнения:
'   Лист MRS очищается от импортированных данных и остаётся готовым к новому заполнению.
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
    ' Сброс MRS использует тот же тихий шаблон восстановления защиты,
    ' что и остальные финальные блоки модуля.
    ProtectSheetQuiet wsIn
    ProtectSheetQuiet wsMRS
    Exit Sub
EH:
    Resume Cleanup
End Sub

' ===========================================================================
' ClearMRSArea
' ===========================================================================
' Назначение:
'   Удаляет только рабочую область данных листа MRS.
'   Верхняя часть листа с шапкой и служебной разметкой при этом сохраняется.
' Параметры:
'   wsMRS - лист MRS.
'   clrLocked - параметр оставлен для единообразия вызова; в текущей реализации напрямую не используется.
' Эффект выполнения:
'   Все строки данных MRS, начиная с четвёртой, удаляются.
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

' ===========================================================================
' SelectAllBrigades
' ===========================================================================
' Назначение:
'   Выбирает все элементы списка на активном DialogSheet.
'   Используется в диалогах выбора дат и бригад при импорте MRS.
' Эффект выполнения:
'   Все элементы первого списка на активном диалоговом листе помечаются как выбранные.
Public Sub SelectAllBrigades()
    On Error Resume Next
    Dim dlg As Object
    Set dlg = ActiveSheet
    If TypeName(dlg) = "DialogSheet" Then
        Dim lb As Object
        Set lb = dlg.ListBoxes(1)
        Dim i As Long
        For i = 1 To lb.ListCount
            lb.selected(i) = True
        Next i
    End If
End Sub

' ===========================================================================
' ClearAllBrigades
' ===========================================================================
' Назначение:
'   Снимает выбор со всех элементов списка на активном DialogSheet.
'   Является парной процедурой к `SelectAllBrigades`.
' Эффект выполнения:
'   Первый список на активном диалоговом листе полностью очищается от выбора.
Public Sub ClearAllBrigades()
    On Error Resume Next
    Dim dlg As Object
    Set dlg = ActiveSheet
    If TypeName(dlg) = "DialogSheet" Then
        Dim lb As Object
        Set lb = dlg.ListBoxes(1)
        Dim i As Long
        For i = 1 To lb.ListCount
            lb.selected(i) = False
        Next i
    End If
End Sub

' ===========================================================================
' SaveHistorySheet
' ===========================================================================
' Назначение:
'   Пользовательская точка входа для экспорта листа истории.
' Эффект выполнения:
'   Передаёт в общую процедуру экспорта лист истории и префикс имени файла.
Public Sub SaveHistorySheet()
    ExportSheet ThisWorkbook.Worksheets(4), UW(1048, 1089, 1090, 1086, 1088, 1080, 1103, 95, 1056, 1072, 1089, 1095, 1077, 1090, 1086, 1074)
End Sub

' ===========================================================================
' SaveMRSSheet
' ===========================================================================
' Назначение:
'   Пользовательская точка входа для экспорта листа MRS.
' Эффект выполнения:
'   Передаёт в общую процедуру экспорта лист MRS и префикс имени файла.
Public Sub SaveMRSSheet()
    ExportSheet ThisWorkbook.Worksheets(5), UW(1055, 1072, 1088, 1089, 1080, 1085, 1075, 95, 77, 82, 83)
End Sub

' ===========================================================================
' ExportSheet
' ===========================================================================
' Назначение:
'   Экспортирует выбранный лист в отдельный файл формата xlsx.
'   Процедура используется и для истории, и для листа MRS, сохраняя текущую
'   структуру листа, но убирая интерактивные элементы исходной макросной книги.
' Основные действия:
'   - Формирует имя файла по умолчанию и показывает диалог Save As.
'   - Копирует выбранный лист в новую временную книгу.
'   - Удаляет с экспортируемого листа кнопки и другие shape-объекты.
'   - Очищает служебную вторую строку.
'   - Повторно создаёт условное форматирование строк заказов по статусу.
'   - Сохраняет новую книгу как обычный xlsx и закрывает её.
' Параметры:
'   ws - лист, который нужно экспортировать.
'   defaultNamePrefix - префикс имени файла по умолчанию.
' Эффект выполнения:
'   Пользователь получает внешний xlsx-файл, подготовленный к открытию без макросной оболочки.
Private Sub ExportSheet(ByVal ws As Worksheet, ByVal defaultNamePrefix As String)
    On Error GoTo EH
    Dim savePath As Variant
    Dim defaultName As String
    Dim newWb As Workbook
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
    
    ws.Copy
    Set newWb = ActiveWorkbook
    Set wsDataExport = newWb.Worksheets(1)
    
    wsDataExport.Unprotect UW(49, 49, 52, 55, 48, 57)

    For Each shp In wsDataExport.Shapes
        shp.Delete
    Next shp

    wsDataExport.Rows(2).ClearContents

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

    wsDataExport.Protect UW(49, 49, 52, 55, 48, 57), True, True, False, False
    wsDataExport.Activate
    
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


