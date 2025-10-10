' ReportByDepartments (final).vb — VB.NET VSTO module
' Features:
'  - Option Strict Off (no late binding)
'  - Two modes:
'      1) Sheets mode: one workbook with sheets per department
'         GenerateFromActiveWorkbook / GenerateFromFile  → return saved .xlsx path
'      2) Files-per-dept mode: one workbook per department into <src>_по_отделам\
'         GenerateFromActiveWorkbookPerDept / GenerateFromFilePerDept → return list of saved paths
'  - Safe AutoFilter (optional, skipped on failure)
'  - Beautify: remove weekend rows with 0:00, highlight 0:00 (G) pale yellow and row font red
'  - Column widths cloned, sheet names/file names sanitized

Option Strict Off
Option Explicit On
Option Infer On

Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Drawing
Imports System.Globalization
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Public Module ReportByDepartments

    ' ====================== КОНСТАНТЫ СТРОК ======================
    Private Const ROW_HEADER As Integer = 5
    Private Const ROW_DATA_START As Integer = 6

    ' ====================== КОНСТАНТЫ КОЛОНОК ======================
    Private Const COL_MARKER As Integer = 1        ' A — "ИТОГО"
    Private Const COL_DEPT As Integer = 2          ' B — отдел (первая строка блока)
    Private Const COL_EMPLOYEE As Integer = 3      ' C — ФИО сотрудника
    Private Const COL_TAB_NUMBER As Integer = 5    ' E — табельный номер
    Private Const COL_DATE As Integer = 6          ' F — дата
    Private Const COL_TIME As Integer = 7          ' G — время
    Private Const COL_ABSENT As Integer = 8        ' H — "Прогулял"
    Private Const COL_SOFT_ABSENT As Integer = 9   ' I — "Мягкие прогулы"
    Private Const COL_START_TIME As Integer = 10   ' J — "Начало дня"
    Private Const COL_END_TIME As Integer = 11     ' K — "Конец дня"
    Private Const COL_HOLIDAY_WORK As Integer = 12 ' L — "Работа в праздничные дни"
    Private Const COL_OVERTIME As Integer = 13     ' M — "Фактическая переработка"
    Private Const COL_WORK_SCHEDULE As Integer = 15 ' O — "График работы"

    ' ====================== КОНСТАНТЫ ФОРМАТИРОВАНИЯ ======================
    Private Const MAX_SHEET_NAME_LENGTH As Integer = 18
    Private Const MAX_FILE_NAME_LENGTH As Integer = 50
    Private Const MAX_SHEET_NAME_FULL_LENGTH As Integer = 31

    ' ====================== ЦВЕТА ======================
    Private Const COLOR_PALE_YELLOW As Integer = &HCCFFFF  ' #FFFFCC (BGR формат)
    Private Const COLOR_LIGHT_RED As Integer = &HC8C8FF    ' #FFC8C8 (BGR формат)
    Private Const COLOR_LIME_GREEN As Integer = &HFF00   ' #00FF00 (BGR формат)
    Private Const COLOR_YELLOW As Integer = &HFFFF       ' #FFFF00 (BGR формат)
    Private Const COLOR_RED As Integer = &HFF          ' #FF0000 (BGR формат)

    ' ====================== СТАНДАРТНЫЕ ЗНАЧЕНИЯ ======================
    Private Const DEFAULT_WORK_SCHEDULE As String = "рабочий график: 8:00-17:00"
    Private Const DEFAULT_SHEET_NAME As String = "Лист"
    Private Const DEFAULT_FILE_NAME As String = "Отчет"
    Private Const TOTAL_MARKER As String = "ИТОГО"
    Private Const NO_ENTRY_TEXT As String = "Нет входа"
    Private Const NO_EXIT_TEXT As String = "Нет выход"

    ' ====================== КОНСТАНТЫ ДЛЯ ПРИЧИН ОТСУТСТВИЯ ======================
    Public Const LEAVE_REASON_HEADER As String = "Причина отсутствия"
    Public Const LEAVES_FILE_HEADER_ROW As Integer = 1
    Public Const LEAVES_FILE_DATA_START_ROW As Integer = 2
    Public Const LEAVES_FILE_COL_EMPLOYEE As Integer = 1    ' A — ФИО сотрудника
    Public Const LEAVES_FILE_COL_START_DATE As Integer = 5  ' E — дата начала отпуска
    Public Const LEAVES_FILE_COL_END_DATE As Integer = 6    ' F — дата окончания отпуска
    Public Const LEAVES_FILE_COL_REASON As Integer = 7      ' G — причина отсутствия

    ' ====================== КОНСТАНТЫ ДЛЯ ГРАФИКА РАБОТЫ ======================
    Public Const WORK_SCHEDULE_HEADER As String = "График работы"
    Public Const SCHEDULES_FILE_HEADER_ROW As Integer = 5
    Public Const SCHEDULES_FILE_DATA_START_ROW As Integer = 7
    Public Const SCHEDULES_FILE_COL_EMPLOYEE As Integer = 3    ' C — ФИО сотрудника
    Public Const SCHEDULES_FILE_COL_SCHEDULE As Integer = 6    ' F — график работы
    Public Const SCHEDULES_FILE_SEARCH_COL As Integer = 3      ' C — колонка для поиска последней строки

    ' ====================== ВСПОМОГАТЕЛЬНЫЕ КЛАССЫ ======================

    ''' <summary>
    ''' Управляет настройками Excel приложения
    ''' </summary>
    Private Class ExcelApplicationManager
        Implements IDisposable

        Private ReadOnly _app As Excel.Application
        Private ReadOnly _originalCalculation As Excel.XlCalculation
        Private ReadOnly _originalScreenUpdating As Boolean
        Private ReadOnly _originalEnableEvents As Boolean
        Private ReadOnly _originalDisplayAlerts As Boolean
        Private _disposed As Boolean = False

        Public Sub New(app As Excel.Application)
            _app = app
            _originalCalculation = app.Calculation
            _originalScreenUpdating = app.ScreenUpdating
            _originalEnableEvents = app.EnableEvents
            _originalDisplayAlerts = app.DisplayAlerts

            ' Устанавливаем оптимальные настройки для работы
            app.Calculation = Excel.XlCalculation.xlCalculationManual
            app.ScreenUpdating = False
            app.EnableEvents = False
            app.DisplayAlerts = False
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If Not _disposed Then
                ' Восстанавливаем оригинальные настройки
                _app.Calculation = _originalCalculation
                _app.ScreenUpdating = _originalScreenUpdating
                _app.EnableEvents = _originalEnableEvents
                _app.DisplayAlerts = _originalDisplayAlerts
                _disposed = True
            End If
        End Sub
    End Class

    ''' <summary>
    ''' Утилиты для работы с Excel объектами
    ''' </summary>
    Private Class ExcelUtilities

        ''' <summary>
        ''' Безопасно получает значение ячейки
        ''' </summary>
        Public Shared Function GetCellValue(ws As Excel.Worksheet, row As Integer, col As Integer) As Object
            Dim rng As Excel.Range = CType(ws.Cells(row, col), Excel.Range)
            Dim v As Object = rng.Value2
            Marshal.FinalReleaseComObject(rng)
            Return v
        End Function

        ''' <summary>
        ''' Безопасно освобождает COM объект
        ''' </summary>
        Public Shared Sub ReleaseComObject(obj As Object)
            If obj IsNot Nothing Then
                Marshal.FinalReleaseComObject(obj)
            End If
        End Sub

        ''' <summary>
        ''' Проверяет, является ли значение строкой "ИТОГО"
        ''' </summary>
        Public Shared Function IsTotalMarker(value As Object) As Boolean
            If value Is Nothing Then Return False
            Dim s As String = CStr(value)
            Return String.Compare(s, TOTAL_MARKER, True, CultureInfo.CurrentCulture) = 0
        End Function

        ''' <summary>
        ''' Проверяет, является ли время нулевым
        ''' </summary>
        Public Shared Function IsZeroTime(value As Object) As Boolean
            If value Is Nothing Then Return False
            If TypeOf value Is Double Then
                Return Math.Abs(CDbl(value)) < 0.0000001R
            End If
            Dim s As String = CStr(value).Trim()
            Return s.StartsWith("0:00", StringComparison.CurrentCulture)
        End Function

        ''' <summary>
        ''' Проверяет, является ли дата выходным днем
        ''' </summary>
        Public Shared Function IsWeekend(value As Object) As Boolean
            If value Is Nothing Then Return False
            Dim dt As Date
            If TypeOf value Is Double Then
                dt = Date.FromOADate(CDbl(value))
            ElseIf Not Date.TryParse(CStr(value), dt) Then
                Return False
            End If
            Dim mondayBased As Integer = ((CInt(dt.DayOfWeek) + 6) Mod 7) + 1 ' 1=Mon .. 7=Sun
            Return (mondayBased = 6 OrElse mondayBased = 7)
        End Function

        ''' <summary>
        ''' Ограничивает имя листа согласно правилам Excel
        ''' </summary>
        Public Shared Function LimitSheetName(proposed As String) As String
            Dim name As String = If(proposed, String.Empty).Trim()
            If name.Length > MAX_SHEET_NAME_LENGTH Then name = name.Substring(0, MAX_SHEET_NAME_LENGTH)
            Dim banned As String = "\/?*[]:"
            For Each ch As Char In banned
                name = name.Replace(ch.ToString(), "_")
            Next
            While name.EndsWith(" ") OrElse name.EndsWith(".")
                name = If(name.Length > 1, name.Substring(0, name.Length - 1), String.Empty)
                If name.Length = 0 Then Exit While
            End While
            If String.IsNullOrWhiteSpace(name) Then name = DEFAULT_SHEET_NAME
            Return name
        End Function

        ''' <summary>
        ''' Ограничивает имя файла согласно правилам Windows
        ''' </summary>
        Public Shared Function LimitFileBaseName(proposed As String) As String
            Dim name As String = If(proposed, String.Empty).Trim()
            If name.Length > MAX_FILE_NAME_LENGTH Then name = name.Substring(0, MAX_FILE_NAME_LENGTH)
            Dim banned As String = "\/:*?""<>|"
            For Each ch As Char In banned
                name = name.Replace(ch.ToString(), "_")
            Next
            While name.EndsWith(" ") OrElse name.EndsWith(".")
                name = If(name.Length > 1, name.Substring(0, name.Length - 1), String.Empty)
                If name.Length = 0 Then Exit While
            End While
            If String.IsNullOrWhiteSpace(name) Then name = DEFAULT_FILE_NAME
            Return name
        End Function

        ''' <summary>
        ''' Создает уникальное имя листа
        ''' </summary>
        Public Shared Function MakeUniqueSheetName(baseName As String, index As Integer) As String
            Dim core As String = baseName
            Dim suffix As String = " (" & index.ToString(CultureInfo.CurrentCulture) & ")"
            If core.Length + suffix.Length > MAX_SHEET_NAME_FULL_LENGTH Then
                core = core.Substring(0, Math.Max(0, MAX_SHEET_NAME_FULL_LENGTH - suffix.Length))
            End If
            Return core & suffix
        End Function

        ''' <summary>
        ''' Пытается найти лист по имени
        ''' </summary>
        Public Shared Function TryGetSheet(wb As Excel.Workbook, name As String) As Excel.Worksheet
            For Each sh As Object In wb.Sheets
                Dim ws = TryCast(sh, Excel.Worksheet)
                If ws IsNot Nothing AndAlso String.Equals(ws.Name, name, StringComparison.CurrentCulture) Then
                    Return ws
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Создает или получает лист с уникальным именем
        ''' </summary>
        Public Shared Function CreateOrGetSheet(wb As Excel.Workbook, baseName As String) As Excel.Worksheet
            Dim nameToUse As String = baseName
            Dim attempt As Integer = 1
            While True
                Dim ws As Excel.Worksheet = TryGetSheet(wb, nameToUse)
                If ws IsNot Nothing Then Return ws
                Try
                    Dim created As Excel.Worksheet = CType(wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)), Excel.Worksheet)
                    created.Name = nameToUse
                    Return created
                Catch ex As Exception
                    attempt += 1
                    nameToUse = MakeUniqueSheetName(baseName, attempt)
                End Try
            End While
        End Function
    End Class

    ''' <summary>
    ''' Управляет листами Excel
    ''' </summary>
    Private Class SheetManager

        ''' <summary>
        ''' Удаляет пустые листы "Отчет"
        ''' </summary>
        Public Shared Sub RemoveEmptyReportSheets(wb As Excel.Workbook)
            Dim sheetsToDelete As New List(Of Excel.Worksheet)

            For Each sh As Object In wb.Sheets
                Dim ws = TryCast(sh, Excel.Worksheet)
                If ws IsNot Nothing AndAlso ws.Name = "Отчет" Then
                    ' Проверяем, пустой ли лист (только заголовки или вообще пустой)
                    Dim usedRange As Excel.Range = ws.UsedRange
                    Dim isEmpty As Boolean = False

                    If usedRange Is Nothing Then
                        isEmpty = True
                    Else
                        Dim rowCount As Integer = usedRange.Rows.Count
                        Dim colCount As Integer = usedRange.Columns.Count
                        isEmpty = (rowCount <= 1 AndAlso colCount <= 1)
                    End If

                    ExcelUtilities.ReleaseComObject(usedRange)

                    If isEmpty Then
                        sheetsToDelete.Add(ws)
                    End If
                End If
            Next

            ' Удаляем найденные пустые листы
            For Each ws As Excel.Worksheet In sheetsToDelete
                ws.Delete()
                ExcelUtilities.ReleaseComObject(ws)
            Next
        End Sub

        ''' <summary>
        ''' Сортирует листы по алфавиту
        ''' </summary>
        Public Shared Sub SortSheetsAlphabetically(wb As Excel.Workbook)
            Dim list As New List(Of Excel.Worksheet)
            For Each sh As Object In wb.Sheets
                Dim ws = TryCast(sh, Excel.Worksheet)
                If ws IsNot Nothing Then list.Add(ws)
            Next

            ' Сортируем так, чтобы листы с "_нет_прохода" были в конце
            list.Sort(Function(a, b)
                          Dim aHasNoPass = a.Name.Contains("_нет_прохода")
                          Dim bHasNoPass = b.Name.Contains("_нет_прохода")

                          ' Если один имеет "_нет_прохода", а другой нет - тот что без суффикса идет первым
                          If aHasNoPass AndAlso Not bHasNoPass Then Return 1
                          If Not aHasNoPass AndAlso bHasNoPass Then Return -1

                          ' Если оба имеют или не имеют "_нет_прохода" - сортируем по алфавиту
                          Return String.Compare(a.Name, b.Name, StringComparison.CurrentCulture)
                      End Function)

            For i As Integer = 0 To list.Count - 1
                Dim ws As Excel.Worksheet = list(i)
                ws.Move(After:=wb.Sheets(i + 1))
            Next
        End Sub

        ''' <summary>
        ''' Копирует заголовок на лист
        ''' </summary>
        Public Shared Sub CopyHeaderToSheet(sourceWs As Excel.Worksheet, targetWs As Excel.Worksheet)
            Dim headerRowRange As Excel.Range = CType(sourceWs.Rows(ROW_HEADER), Excel.Range)
            Dim destHeaderRow As Excel.Range = CType(targetWs.Rows(1), Excel.Range)
            headerRowRange.Copy(Destination:=destHeaderRow)
            destHeaderRow.Font.Bold = True
            ExcelUtilities.ReleaseComObject(headerRowRange)
            ExcelUtilities.ReleaseComObject(destHeaderRow)
        End Sub

        ''' <summary>
        ''' Копирует диапазон строк на лист
        ''' </summary>
        Public Shared Sub CopyRowsToSheet(sourceWs As Excel.Worksheet, targetWs As Excel.Worksheet, startRow As Integer, endRow As Integer)
            Dim pasteRow As Integer = targetWs.Cells(targetWs.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row + 1
            Dim srcRange As Excel.Range = sourceWs.Range(sourceWs.Rows(startRow), sourceWs.Rows(endRow))
            Dim destPaste As Excel.Range = CType(targetWs.Rows(pasteRow), Excel.Range)
            srcRange.Copy(Destination:=destPaste)
            ExcelUtilities.ReleaseComObject(srcRange)
            ExcelUtilities.ReleaseComObject(destPaste)
        End Sub

        ''' <summary>
        ''' Применяет автофильтр к листу
        ''' </summary>
        Public Shared Sub ApplyAutoFilter(ws As Excel.Worksheet, headerRow As Integer)
            Try
                If ws Is Nothing Then Exit Sub
                Dim lastRow As Integer = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
                Dim lastCol As Integer = ws.Cells(headerRow, ws.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
                If lastRow < headerRow OrElse lastCol < 1 Then Exit Sub
                ws.Activate()
                If ws.AutoFilterMode Then ws.AutoFilterMode = False
                Dim rng As Excel.Range = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))
                Try
                    rng.AutoFilter()
                Catch ex As COMException
                    Dim ur As Excel.Range = ws.UsedRange
                    Try
                        ur.AutoFilter()
                    Finally
                        ExcelUtilities.ReleaseComObject(ur)
                    End Try
                Finally
                    ExcelUtilities.ReleaseComObject(rng)
                End Try
            Catch
                ' ignore
            End Try
        End Sub
    End Class

    ''' <summary>
    ''' Управляет колонками Excel
    ''' </summary>
    Private Class ColumnManager

        ''' <summary>
        ''' Удаляет ненужные колонки из листа и переносит "ИТОГО" во вторую колонку
        ''' </summary>
        Public Shared Sub RemoveUnnecessaryColumns(ws As Excel.Worksheet)
            Try
                ' Переименовываем колонку "Мягкие прогулы" в "Находился вне здания" и добавляем комментарий
                RenameSoftAbsentColumn(ws)

                ' Переносим "ИТОГО" из первой колонки во вторую
                MoveTotalMarkersToSecondColumn(ws)

                ' Удаляем ненужные колонки в обратном порядке
                DeleteUnnecessaryColumns(ws)

            Catch ex As Exception
                ' Игнорируем ошибки удаления колонок
            End Try
        End Sub

        ''' <summary>
        ''' Переименовывает колонку "Мягкие прогулы" и добавляет комментарий
        ''' </summary>
        Private Shared Sub RenameSoftAbsentColumn(ws As Excel.Worksheet)
            Try
                Dim headerCell As Excel.Range = CType(ws.Cells(1, COL_SOFT_ABSENT), Excel.Range)
                Dim headerValue As Object = ExcelUtilities.GetCellValue(ws, 1, COL_SOFT_ABSENT)

                If Not IsNothing(headerValue) AndAlso CStr(headerValue).Contains("Мягких прогулов") Then
                    ' Переименовываем заголовок
                    headerCell.Value2 = "Находился вне здания"

                    ' Добавляем комментарий
                    Dim commentText As String = "Время, которое сотрудник находился вне здания в необеденное время"
                    If headerCell.Comment IsNot Nothing Then
                        headerCell.Comment.Delete()
                    End If
                    headerCell.AddComment(commentText)

                    ' Настраиваем размер комментария
                    If headerCell.Comment IsNot Nothing Then
                        headerCell.Comment.Shape.Width = 400
                        headerCell.Comment.Shape.Height = 150
                        headerCell.Comment.Shape.TextFrame.AutoSize = True

                        ' Увеличиваем padding (отступы) для комментария
                        With headerCell.Comment.Shape.TextFrame
                            .MarginLeft = 10
                            .MarginRight = 10
                            .MarginTop = 10
                            .MarginBottom = 10
                        End With
                    End If

                    ' Выравниваем ширину столбца
                    Dim columnRange As Excel.Range = CType(ws.Columns(COL_SOFT_ABSENT), Excel.Range)
                    columnRange.AutoFit()
                    ExcelUtilities.ReleaseComObject(columnRange)
                End If

                ExcelUtilities.ReleaseComObject(headerCell)
            Catch ex As Exception
                ' Игнорируем ошибки переименования
            End Try
        End Sub

        ''' <summary>
        ''' Переносит маркеры "ИТОГО" из первой колонки во вторую
        ''' </summary>
        Private Shared Sub MoveTotalMarkersToSecondColumn(ws As Excel.Worksheet)
            Dim lastRow As Integer = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            For r As Integer = ROW_DATA_START To lastRow
                Dim markerObj As Object = ExcelUtilities.GetCellValue(ws, r, 1) ' Первая колонка
                If Not IsNothing(markerObj) AndAlso String.Equals(CStr(markerObj), TOTAL_MARKER, StringComparison.CurrentCultureIgnoreCase) Then
                    ' Копируем "ИТОГО" во вторую колонку
                    Dim targetCell As Excel.Range = CType(ws.Cells(r, 2), Excel.Range)
                    targetCell.Value2 = TOTAL_MARKER
                    ExcelUtilities.ReleaseComObject(targetCell)

                    ' Очищаем первую колонку
                    Dim sourceCell As Excel.Range = CType(ws.Cells(r, 1), Excel.Range)
                    sourceCell.Value2 = ""
                    ExcelUtilities.ReleaseComObject(sourceCell)
                End If
            Next
        End Sub

        ''' <summary>
        ''' Удаляет ненужные колонки
        ''' </summary>
        Private Shared Sub DeleteUnnecessaryColumns(ws As Excel.Worksheet)
            ' Удаляем "Работа в праздничные дни" (12-я колонка)
            Dim holidayCol As Excel.Range = CType(ws.Columns(COL_HOLIDAY_WORK), Excel.Range)
            holidayCol.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
            ExcelUtilities.ReleaseComObject(holidayCol)

            ' Удаляем "Прогулял" (8-я колонка)
            Dim absentCol As Excel.Range = CType(ws.Columns(COL_ABSENT), Excel.Range)
            absentCol.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
            ExcelUtilities.ReleaseComObject(absentCol)

            ' Удаляем "Таб #" (5-я колонка)
            Dim tabCol As Excel.Range = CType(ws.Columns(COL_TAB_NUMBER), Excel.Range)
            tabCol.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
            ExcelUtilities.ReleaseComObject(tabCol)

            ' Удаляем первую колонку (теперь пустую)
            Dim firstCol As Excel.Range = CType(ws.Columns(1), Excel.Range)
            firstCol.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
            ExcelUtilities.ReleaseComObject(firstCol)
        End Sub
    End Class

    ' ====================== SHEETS MODE (one workbook with many sheets) ======================
    Public Function GenerateFromActiveWorkbook(app As Excel.Application) As String
        If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))
        Dim wb As Excel.Workbook = app.ActiveWorkbook
        If wb Is Nothing Then Throw New InvalidOperationException("Нет активной книги.")
        Return GenerateCore(app, wb)
    End Function

    Public Function GenerateFromFile(app As Excel.Application, workbookPath As String) As String
        If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))
        If String.IsNullOrWhiteSpace(workbookPath) OrElse Not File.Exists(workbookPath) Then
            Throw New FileNotFoundException("Файл не найден", workbookPath)
        End If
        Dim opened As Boolean = False
        Dim wb As Excel.Workbook = Nothing
        Try
            wb = app.Workbooks.Open(workbookPath, ReadOnly:=False)
            opened = True
            Return GenerateCore(app, wb)
        Finally
            If opened AndAlso wb IsNot Nothing Then
                wb.Close(SaveChanges:=False)
                Marshal.FinalReleaseComObject(wb)
            End If
        End Try
    End Function

    Private Function GenerateCore(app As Excel.Application, srcWb As Excel.Workbook) As String
        Dim savedPath As String = Nothing

        Using appManager As New ExcelApplicationManager(app)
            Try
                ' Получаем исходный лист
                Dim wsSource = GetSourceWorksheet(app, srcWb)
                SheetManager.ApplyAutoFilter(wsSource, ROW_HEADER)

                ' Создаем выходную книгу
                Dim wbNew As Excel.Workbook = CreateOutputWorkbook(app, srcWb)
                Dim savePath As String = SaveWorkbook(wbNew, srcWb, app)

                ' Обрабатываем данные по отделам
                ProcessDepartmentData(wsSource, wbNew)

                ' Финальная обработка листов
                FinalizeWorkbook(wbNew)

                app.StatusBar = $"Готово! Файл сохранён: {Path.GetFileName(savePath)}"
                srcWb.Activate()
                savedPath = savePath
            Catch ex As Exception
                ' Логируем ошибку, если необходимо
                Throw
            End Try
        End Using

        Return savedPath
    End Function

    ''' <summary>
    ''' Получает исходный рабочий лист
    ''' </summary>
    Private Function GetSourceWorksheet(app As Excel.Application, srcWb As Excel.Workbook) As Excel.Worksheet
        Dim wsSource = TryCast(app.ActiveSheet, Excel.Worksheet)
        If wsSource Is Nothing OrElse Not Object.ReferenceEquals(wsSource.Parent, srcWb) Then
            wsSource = CType(srcWb.Sheets(1), Excel.Worksheet)
            wsSource.Activate()
        End If
        Return wsSource
    End Function

    ''' <summary>
    ''' Создает выходную рабочую книгу
    ''' </summary>
    Private Function CreateOutputWorkbook(app As Excel.Application, srcWb As Excel.Workbook) As Excel.Workbook
        Dim wbNew As Excel.Workbook = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet)
        Dim wsFirst As Excel.Worksheet = CType(wbNew.Sheets(1), Excel.Worksheet)
        wsFirst.Cells.Clear()
        wsFirst.Name = "Отчет"
        Return wbNew
    End Function

    ''' <summary>
    ''' Сохраняет рабочую книгу
    ''' </summary>
    Private Function SaveWorkbook(wb As Excel.Workbook, srcWb As Excel.Workbook, app As Excel.Application) As String
        Dim baseName As String = Path.GetFileNameWithoutExtension(srcWb.Name)
        Dim newName As String = baseName & "_по_отделам.xlsx"
        Dim saveDir As String = If(String.IsNullOrEmpty(srcWb.Path), app.DefaultFilePath, srcWb.Path)
        Dim savePath As String = Path.Combine(saveDir, newName)
        If File.Exists(savePath) Then File.Delete(savePath)
        wb.SaveAs(Filename:=savePath, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook)
        Return savePath
    End Function

    ''' <summary>
    ''' Обрабатывает данные по отделам
    ''' </summary>
    Private Sub ProcessDepartmentData(wsSource As Excel.Worksheet, wbNew As Excel.Workbook)
        Dim lastRow As Integer = wsSource.Cells(wsSource.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        Dim created As New Dictionary(Of String, Excel.Worksheet)(StringComparer.CurrentCulture)
        Dim startCopyRow As Integer = ROW_DATA_START

        For r As Integer = ROW_HEADER + 1 To lastRow
            Dim markerObj As Object = ExcelUtilities.GetCellValue(wsSource, r, COL_MARKER)
            If ExcelUtilities.IsTotalMarker(markerObj) Then
                ProcessDepartmentBlock(wsSource, wbNew, created, startCopyRow, r)
                startCopyRow = r + 1
            End If
        Next
    End Sub

    ''' <summary>
    ''' Обрабатывает блок данных отдела
    ''' </summary>
    Private Sub ProcessDepartmentBlock(wsSource As Excel.Worksheet, wbNew As Excel.Workbook,
                                     created As Dictionary(Of String, Excel.Worksheet),
                                     startCopyRow As Integer, endRow As Integer)
        Dim deptRaw As String = CStr(ExcelUtilities.GetCellValue(wsSource, startCopyRow, COL_DEPT))
        Dim dept As String = ExcelUtilities.LimitSheetName(deptRaw)
        Dim timeObj As Object = ExcelUtilities.GetCellValue(wsSource, endRow, COL_TIME)
        If ExcelUtilities.IsZeroTime(timeObj) Then dept &= "_нет_прохода"

        If dept.Length > 0 Then
            Dim wsTarget As Excel.Worksheet = GetOrCreateDepartmentSheet(wbNew, created, dept, wsSource)
            SheetManager.CopyRowsToSheet(wsSource, wsTarget, startCopyRow, endRow)
            BeautifySheet(wsTarget)
            AutoFitColumns(wsTarget)
        End If
    End Sub

    ''' <summary>
    ''' Получает или создает лист отдела
    ''' </summary>
    Private Function GetOrCreateDepartmentSheet(wb As Excel.Workbook, created As Dictionary(Of String, Excel.Worksheet),
                                              dept As String, wsSource As Excel.Worksheet) As Excel.Worksheet
        Dim wsTarget As Excel.Worksheet = Nothing
        If Not created.TryGetValue(dept, wsTarget) OrElse wsTarget Is Nothing Then
            wsTarget = ExcelUtilities.CreateOrGetSheet(wb, dept)
            SheetManager.CopyHeaderToSheet(wsSource, wsTarget)
            created(dept) = wsTarget
        End If
        Return wsTarget
    End Function

    ''' <summary>
    ''' Автоматически подгоняет ширину колонок
    ''' </summary>
    Private Sub AutoFitColumns(ws As Excel.Worksheet)
        Dim rngFit As Excel.Range = ws.UsedRange
        rngFit.Columns.AutoFit()
        ExcelUtilities.ReleaseComObject(rngFit)
    End Sub

    ''' <summary>
    ''' Финальная обработка рабочей книги
    ''' </summary>
    Private Sub FinalizeWorkbook(wb As Excel.Workbook)
        ' Удаляем ненужные колонки из всех листов
        For Each ws As Excel.Worksheet In wb.Sheets
            ColumnManager.RemoveUnnecessaryColumns(ws)
        Next

        ' Удаляем пустые листы "Отчет"
        SheetManager.RemoveEmptyReportSheets(wb)

        ' Сортируем листы по алфавиту
        SheetManager.SortSheetsAlphabetically(wb)
    End Sub

    ' ====================== FILES-PER-DEPT MODE (one workbook per department) ======================
    Public Function GenerateFromActiveWorkbookPerDept(app As Excel.Application) As List(Of String)
        If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))
        Dim wb As Excel.Workbook = app.ActiveWorkbook
        If wb Is Nothing Then Throw New InvalidOperationException("Нет активной книги.")
        Return GenerateCorePerDept(app, wb)
    End Function

    Public Function GenerateFromFilePerDept(app As Excel.Application, workbookPath As String) As List(Of String)
        If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))
        If String.IsNullOrWhiteSpace(workbookPath) OrElse Not File.Exists(workbookPath) Then
            Throw New FileNotFoundException("Файл не найден", workbookPath)
        End If
        Dim opened As Boolean = False
        Dim wb As Excel.Workbook = Nothing
        Try
            wb = app.Workbooks.Open(workbookPath, ReadOnly:=False)
            opened = True
            Return GenerateCorePerDept(app, wb)
        Finally
            If opened AndAlso wb IsNot Nothing Then
                wb.Close(SaveChanges:=False)
                Marshal.FinalReleaseComObject(wb)
            End If
        End Try
    End Function

    Private Function GenerateCorePerDept(app As Excel.Application, srcWb As Excel.Workbook) As List(Of String)
        Dim savedPaths As New List(Of String)()

        Using appManager As New ExcelApplicationManager(app)
            Try
                ' Получаем исходный лист
                Dim wsSource = GetSourceWorksheet(app, srcWb)
                SheetManager.ApplyAutoFilter(wsSource, ROW_HEADER)

                ' Создаем целевую директорию
                Dim targetDir As String = CreateTargetDirectory(srcWb, app)

                ' Обрабатываем данные по отделам в отдельные файлы
                ProcessDepartmentDataPerDept(wsSource, app, targetDir, savedPaths)

                app.StatusBar = $"Готово! Создано файлов: {savedPaths.Count}. Папка: {targetDir}"
                srcWb.Activate()
            Catch ex As Exception
                ' Логируем ошибку, если необходимо
                Throw
            End Try
        End Using

        Return savedPaths
    End Function

    ''' <summary>
    ''' Создает целевую директорию для файлов по отделам
    ''' </summary>
    Private Function CreateTargetDirectory(srcWb As Excel.Workbook, app As Excel.Application) As String
        Dim saveRoot As String = If(String.IsNullOrEmpty(srcWb.Path), app.DefaultFilePath, srcWb.Path)
        Dim baseNamePerDept As String = Path.GetFileNameWithoutExtension(srcWb.Name)
        Dim targetDir As String = Path.Combine(saveRoot, ExcelUtilities.LimitFileBaseName(baseNamePerDept) & "_по_отделам")
        If Not Directory.Exists(targetDir) Then Directory.CreateDirectory(targetDir)
        Return targetDir
    End Function

    ''' <summary>
    ''' Обрабатывает данные по отделам в отдельные файлы
    ''' </summary>
    Private Sub ProcessDepartmentDataPerDept(wsSource As Excel.Worksheet, app As Excel.Application,
                                           targetDir As String, savedPaths As List(Of String))
        Dim lastRow As Integer = wsSource.Cells(wsSource.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        Dim deptToWb As New Dictionary(Of String, Excel.Workbook)(StringComparer.CurrentCulture)
        Dim deptToPath As New Dictionary(Of String, String)(StringComparer.CurrentCulture)
        Dim createdSheets As New Dictionary(Of String, Excel.Worksheet)(StringComparer.CurrentCulture)
        Dim deptHasAnySheet As New HashSet(Of String)(StringComparer.CurrentCulture)
        Dim startCopyRow As Integer = ROW_DATA_START

        For r As Integer = ROW_HEADER + 1 To lastRow
            Dim markerObj As Object = ExcelUtilities.GetCellValue(wsSource, r, COL_MARKER)
            If ExcelUtilities.IsTotalMarker(markerObj) Then
                ProcessDepartmentBlockPerDept(wsSource, app, targetDir, deptToWb, deptToPath,
                                            createdSheets, deptHasAnySheet, savedPaths, startCopyRow, r)
                startCopyRow = r + 1
            End If
        Next

        ' Финальная обработка всех рабочих книг
        FinalizeDepartmentWorkbooks(deptToWb)
    End Sub

    ''' <summary>
    ''' Обрабатывает блок данных отдела для режима "файл на отдел"
    ''' </summary>
    Private Sub ProcessDepartmentBlockPerDept(wsSource As Excel.Worksheet, app As Excel.Application,
                                            targetDir As String, deptToWb As Dictionary(Of String, Excel.Workbook),
                                            deptToPath As Dictionary(Of String, String),
                                            createdSheets As Dictionary(Of String, Excel.Worksheet),
                                            deptHasAnySheet As HashSet(Of String), savedPaths As List(Of String),
                                            startCopyRow As Integer, endRow As Integer)
        Dim deptRaw As String = CStr(ExcelUtilities.GetCellValue(wsSource, startCopyRow, COL_DEPT))
        Dim deptBase As String = ExcelUtilities.LimitSheetName(deptRaw)
        If deptBase.Length = 0 Then Return

        Dim sheetName As String = deptBase
        Dim timeObj As Object = ExcelUtilities.GetCellValue(wsSource, endRow, COL_TIME)
        If ExcelUtilities.IsZeroTime(timeObj) Then sheetName &= "_нет_прохода"

        ' Получаем или создаем рабочую книгу для отдела
        Dim wbDept As Excel.Workbook = GetOrCreateDepartmentWorkbook(app, targetDir, deptBase, deptToWb, deptToPath, savedPaths)

        ' Получаем или создаем лист в рабочей книге отдела
        Dim wsTarget As Excel.Worksheet = GetOrCreateDepartmentSheetPerDept(wbDept, deptBase, sheetName,
                                                                           createdSheets, deptHasAnySheet, wsSource)

        ' Копируем данные
        SheetManager.CopyRowsToSheet(wsSource, wsTarget, startCopyRow, endRow)
        BeautifySheet(wsTarget)
        AutoFitColumns(wsTarget)

        ' Сохраняем рабочую книгу
        wbDept.Save()
    End Sub

    ''' <summary>
    ''' Получает или создает рабочую книгу для отдела
    ''' </summary>
    Private Function GetOrCreateDepartmentWorkbook(app As Excel.Application, targetDir As String,
                                                  deptBase As String, deptToWb As Dictionary(Of String, Excel.Workbook),
                                                  deptToPath As Dictionary(Of String, String),
                                                  savedPaths As List(Of String)) As Excel.Workbook
        Dim wbDept As Excel.Workbook = Nothing
        If Not deptToWb.TryGetValue(deptBase, wbDept) Then
            wbDept = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet)
            CType(wbDept.Sheets(1), Excel.Worksheet).Name = "Отчет"
            Dim wbPath As String = Path.Combine(targetDir, ExcelUtilities.LimitFileBaseName(deptBase) & ".xlsx")
            If File.Exists(wbPath) Then File.Delete(wbPath)
            wbDept.SaveAs(Filename:=wbPath, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook)
            deptToWb(deptBase) = wbDept
            deptToPath(deptBase) = wbPath
            savedPaths.Add(wbPath)
        End If
        Return wbDept
    End Function

    ''' <summary>
    ''' Получает или создает лист в рабочей книге отдела
    ''' </summary>
    Private Function GetOrCreateDepartmentSheetPerDept(wbDept As Excel.Workbook, deptBase As String,
                                                      sheetName As String, createdSheets As Dictionary(Of String, Excel.Worksheet),
                                                      deptHasAnySheet As HashSet(Of String), wsSource As Excel.Worksheet) As Excel.Worksheet
        Dim key As String = deptBase & "|" & sheetName
        Dim wsTarget As Excel.Worksheet = Nothing
        If Not createdSheets.TryGetValue(key, wsTarget) Then
            If Not deptHasAnySheet.Contains(deptBase) AndAlso wbDept.Sheets.Count = 1 AndAlso CType(wbDept.Sheets(1), Excel.Worksheet).Name = "Отчет" Then
                wsTarget = CType(wbDept.Sheets(1), Excel.Worksheet)
                wsTarget.Name = sheetName
            Else
                wsTarget = CType(wbDept.Sheets.Add(After:=wbDept.Sheets(wbDept.Sheets.Count)), Excel.Worksheet)
                wsTarget.Name = sheetName
            End If
            SheetManager.CopyHeaderToSheet(wsSource, wsTarget)
            createdSheets(key) = wsTarget
            deptHasAnySheet.Add(deptBase)
        End If
        Return wsTarget
    End Function

    ''' <summary>
    ''' Финальная обработка рабочих книг отделов
    ''' </summary>
    Private Sub FinalizeDepartmentWorkbooks(deptToWb As Dictionary(Of String, Excel.Workbook))
        For Each kv In deptToWb
            Dim wbDept As Excel.Workbook = kv.Value

            ' Удаляем ненужные колонки из всех листов
            For Each ws As Excel.Worksheet In wbDept.Sheets
                ColumnManager.RemoveUnnecessaryColumns(ws)
            Next

            ' Удаляем пустые листы "Отчет"
            SheetManager.RemoveEmptyReportSheets(wbDept)

            ' Сортируем листы по алфавиту
            SheetManager.SortSheetsAlphabetically(wbDept)

            wbDept.Save()
            wbDept.Close(SaveChanges:=False)
            ExcelUtilities.ReleaseComObject(wbDept)
        Next
    End Sub

    ' ====================== Helpers ======================

    Private Sub BeautifySheet(ws As Excel.Worksheet)
        Dim lastRow As Integer = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row

        ' Сначала удаляем выходные строки
        For r As Integer = lastRow To 2 Step -1
            Dim dateVal As Object = ExcelUtilities.GetCellValue(ws, r, COL_DATE)
            Dim timeVal As Object = ExcelUtilities.GetCellValue(ws, r, COL_TIME)
            'удаляем выходные
            If ExcelUtilities.IsWeekend(dateVal) AndAlso ExcelUtilities.IsZeroTime(timeVal) Then
                CType(ws.Rows(r), Excel.Range).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                Continue For
            End If
        Next

        ' Теперь обрабатываем оставшиеся строки
        lastRow = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row

        ' ПЕРВЫЙ ЭТАП: Обработка всех строк (окрашивание, анализ времени, установка 0)
        For r As Integer = lastRow To 2 Step -1
            Dim dateVal As Object = ExcelUtilities.GetCellValue(ws, r, COL_DATE)
            Dim timeVal As Object = ExcelUtilities.GetCellValue(ws, r, COL_TIME)

            'красим 0 проходы
            If ExcelUtilities.IsZeroTime(timeVal) Then
                Dim tcell As Excel.Range = CType(ws.Cells(r, COL_TIME), Excel.Range)
                tcell.Interior.Color = COLOR_PALE_YELLOW ' фон желтый
                tcell.Font.Color = COLOR_RED ' шрифт красный
                ExcelUtilities.ReleaseComObject(tcell)
            End If

            ' ==================== АНАЛИЗ ОПОЗДАНИЙ И РАННИХ УХОДОВ ====================
            ' Анализируем время прихода/ухода только для рабочих дней и не для строк ИТОГО
            If Not ExcelUtilities.IsWeekend(dateVal) AndAlso Not ExcelUtilities.IsTotalMarker(ws.Cells(r, COL_MARKER).Value2) Then
                AnalyzeWorkTimeViolations(ws, r)

                ' Проверяем, есть ли вход и выход для сотрудника
                Dim startTime As String = CStr(ExcelUtilities.GetCellValue(ws, r, COL_START_TIME))
                Dim endTime As String = CStr(ExcelUtilities.GetCellValue(ws, r, COL_END_TIME))

                ' Если нет входа или выхода, ставим 0 в фактическую переработку
                If String.IsNullOrEmpty(startTime) OrElse String.IsNullOrEmpty(endTime) OrElse
                   startTime.Contains(NO_ENTRY_TEXT) OrElse endTime.Contains(NO_EXIT_TEXT) Then
                    Dim overtimeCell As Excel.Range = CType(ws.Cells(r, COL_OVERTIME), Excel.Range)
                    overtimeCell.Value2 = 0
                    ExcelUtilities.ReleaseComObject(overtimeCell)

                End If
            End If
        Next

        ' ВТОРОЙ ЭТАП: Суммирование для строк ИТОГО
        For r As Integer = lastRow To 2 Step -1

            ' --- Обработка "Фактическая переработка" ТОЛЬКО для строки ИТОГО (13-й столбец) ---
            Dim marker As Object = ws.Cells(r, COL_MARKER).Value2
            If Not IsNothing(marker) AndAlso ExcelUtilities.IsTotalMarker(marker) Then

                ' Добавляем формулу для пересчета фактической переработки по сотруднику
                Dim ocell As Excel.Range = CType(ws.Cells(r, COL_OVERTIME), Excel.Range) ' 13-й столбец

                ' Ищем начало блока сотрудника (предыдущая строка ИТОГО + 1)
                Dim formulaStartRow As Integer = r - 1
                While formulaStartRow > 1 AndAlso Not ExcelUtilities.IsTotalMarker(ws.Cells(formulaStartRow, COL_MARKER).Value2)
                    formulaStartRow -= 1
                End While

                ' Если нашли предыдущую строку ИТОГО, начинаем с следующей строки
                If ExcelUtilities.IsTotalMarker(ws.Cells(formulaStartRow, COL_MARKER).Value2) Then
                    formulaStartRow += 1
                End If

                ' Исправляем логику: если не нашли предыдущий ИТОГО, начинаем с первой строки данных
                If formulaStartRow <= 1 Then
                    formulaStartRow = 2 ' Первая строка данных после заголовка
                End If


                ' Проверяем, что есть строки для суммирования
                If formulaStartRow < r - 1 Then
                    ' Считаем сумму программно
                    Dim totalHours As Double = 0

                    For row As Integer = formulaStartRow To r - 1
                        Dim cellValue As Object = ExcelUtilities.GetCellValue(ws, row, COL_OVERTIME)
                        If cellValue IsNot Nothing Then
                            ' Пробуем преобразовать в число разными способами
                            Dim numericValue As Double = 0
                            If IsNumeric(cellValue) Then
                                numericValue = CDbl(cellValue)
                            Else
                                ' Пробуем преобразовать текст времени в число
                                Dim timeStr As String = cellValue.ToString()
                                If timeStr.Contains(":") Then
                                    Try
                                        ' Парсим время в формате "h:mm" или "-h:mm"
                                        Dim isNegative As Boolean = timeStr.StartsWith("-")
                                        If isNegative Then timeStr = timeStr.Substring(1)

                                        Dim parts() As String = timeStr.Split(":"c)
                                        If parts.Length = 2 Then
                                            Dim timeHours As Integer = Integer.Parse(parts(0))
                                            Dim timeMinutes As Integer = Integer.Parse(parts(1))
                                            numericValue = (timeHours + timeMinutes / 60.0) / 24.0 ' Конвертируем в дни Excel
                                            If isNegative Then numericValue = -numericValue
                                        End If
                                    Catch
                                        ' Игнорируем ошибки парсинга
                                    End Try
                                End If
                            End If
                            totalHours += numericValue
                        End If
                    Next

                    ' Устанавливаем вычисленное значение в текстовом формате
                    Dim absHours As Double = Math.Abs(totalHours)
                    Dim totalMinutes As Integer = CInt(absHours * 24 * 60)
                    Dim resultHours As Integer = totalMinutes \ 60
                    Dim resultMinutes As Integer = totalMinutes Mod 60

                    If totalHours < 0 Then
                        ocell.Value2 = $"-{resultHours}:{resultMinutes:D2}"
                    Else
                        ocell.Value2 = $"{resultHours}:{resultMinutes:D2}"
                    End If
                    ocell.NumberFormat = "@" ' Текстовый формат для всех значений

                Else
                    ' Если нет строк для суммирования, ставим 0
                    ocell.Value2 = "0:00"
                    ocell.NumberFormat = "@" ' Текстовый формат
                End If
                Dim ov As Object = ocell.Value2

                Dim hours As Double
                Dim haveHours As Boolean = False

                If IsNumeric(ov) Then
                    ' Число: либо часы, либо дни (если ячейка в формате времени)
                    Dim v As Double = CDbl(ov)
                    Dim nf As String = CStr(ocell.NumberFormat)
                    Dim isTimeFmt As Boolean = (nf.IndexOf(":", StringComparison.Ordinal) >= 0) OrElse
                                   (nf.IndexOf("[h", StringComparison.OrdinalIgnoreCase) >= 0)
                    hours = If(isTimeFmt, v * 24.0R, v)
                    haveHours = True
                Else
                    ' Текст вроде "-167:00" → распарсим как часы:минуты(:секунды)
                    Dim s As String = CStr(ov).Trim()
                    Dim sign As Double = 1
                    If s.StartsWith("-"c) Then sign = -1 : s = s.Substring(1)
                    If s.StartsWith("+"c) Then s = s.Substring(1)
                    Dim parts() As String = s.Split(":"c)
                    If parts.Length >= 2 Then
                        Dim hh As Double, mm As Double, ss As Double
                        If Double.TryParse(parts(0), hh) AndAlso Double.TryParse(parts(1), mm) Then
                            If parts.Length >= 3 Then Double.TryParse(parts(2), ss)
                            hours = sign * (hh + mm / 60.0R + ss / 3600.0R)
                            haveHours = True
                        End If
                    End If
                End If

                If haveHours Then
                    ' Снимем CF для этой ячейки, чтобы не перебивало заливку
                    ocell.FormatConditions.Delete()

                    If hours > 0 Then
                        ocell.Interior.Color = COLOR_LIME_GREEN        ' > 0 → зелёный
                    ElseIf hours >= -1 Then
                        ocell.Interior.Color = COLOR_YELLOW           ' -1..0 → жёлтый (напр. -0,27)
                    Else
                        ocell.Interior.Color = COLOR_RED              ' < -1 → красный (напр. -1,27 или -167)
                    End If
                End If

                ExcelUtilities.ReleaseComObject(ocell)
            End If
        Next
        ' === Границы: применяем один раз для всего заполненного диапазона ===
        Dim lastRowAll As Integer = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        Dim lastColAll As Integer = ws.Cells(1, ws.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

        If lastRowAll >= 1 AndAlso lastColAll >= 1 Then
            Dim rngAll As Excel.Range = ws.Range(ws.Cells(1, 1), ws.Cells(lastRowAll, lastColAll))
            Dim borders As Excel.Borders = rngAll.Borders

            With borders
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = Excel.XlBorderWeight.xlThin
                .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic
            End With

            ' Гарантируем внутренние линии (иногда коллекционная установка их не трогает)
            borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
            borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous

            ExcelUtilities.ReleaseComObject(borders)
            ExcelUtilities.ReleaseComObject(rngAll)
        End If

    End Sub


    ' ==================== АНАЛИЗ ОПОЗДАНИЙ И РАННИХ УХОДОВ ====================
    ' Анализирует время прихода/ухода и подсвечивает нарушения
    Private Sub AnalyzeWorkTimeViolations(ws As Excel.Worksheet, row As Integer)
        Try
            ' Читаем ФИО сотрудника
            Dim fio As String = CStr(ExcelUtilities.GetCellValue(ws, row, COL_EMPLOYEE))
            If String.IsNullOrEmpty(fio) Then Return

            ' Читаем время прихода и ухода
            Dim startTime As String = CStr(ExcelUtilities.GetCellValue(ws, row, COL_START_TIME))
            Dim endTime As String = CStr(ExcelUtilities.GetCellValue(ws, row, COL_END_TIME))

            ' Пропускаем если нет данных о времени
            If String.IsNullOrEmpty(startTime) AndAlso String.IsNullOrEmpty(endTime) Then Return
            If startTime.Contains(NO_ENTRY_TEXT) AndAlso endTime.Contains(NO_EXIT_TEXT) Then Return

            ' Получаем график работы для сотрудника из колонки графика работы
            Dim workSchedule As String = GetWorkScheduleForEmployee(ws, row)
            If String.IsNullOrEmpty(workSchedule) Then Return

            ' Если график стандартный (не проставлен), не анализируем
            If workSchedule = DEFAULT_WORK_SCHEDULE Then Return

            ' Получаем дату для определения дня недели
            Dim dayValue As Date? = LeaveReasonFiller.ReadDate(ws, row, COL_DATE)
            Dim isFriday As Boolean = dayValue.HasValue AndAlso dayValue.Value.DayOfWeek = DayOfWeek.Friday

            ' Извлекаем время начала и окончания работы из графика
            Dim expectedStart As TimeSpan? = ExtractStartTimeFromSchedule(workSchedule)
            Dim expectedEnd As TimeSpan? = ExtractEndTimeFromSchedule(workSchedule, isFriday)

            ' Получаем объекты ячеек для анализа
            Dim startTimeObj As Object = ExcelUtilities.GetCellValue(ws, row, COL_START_TIME)
            Dim endTimeObj As Object = ExcelUtilities.GetCellValue(ws, row, COL_END_TIME)


            Dim hasViolation As Boolean = False
            Dim violationText As String = ""

            ' Анализ времени прихода
            If Not String.IsNullOrEmpty(startTime) AndAlso Not startTime.Contains(NO_ENTRY_TEXT) AndAlso expectedStart.HasValue Then
                Dim actualStart As TimeSpan? = ParseTimeFromCellValue(startTimeObj)

                If actualStart.HasValue AndAlso actualStart.Value > expectedStart.Value Then
                    Dim delay As TimeSpan = actualStart.Value - expectedStart.Value
                    violationText += $"Опоздание: {delay.Hours}ч {delay.Minutes}м. "
                    hasViolation = True
                End If
            End If

            ' Анализ времени ухода
            If Not String.IsNullOrEmpty(endTime) AndAlso Not endTime.Contains(NO_EXIT_TEXT) AndAlso expectedEnd.HasValue Then
                Dim actualEnd As TimeSpan? = ParseTimeFromCellValue(endTimeObj)

                If actualEnd.HasValue AndAlso actualEnd.Value < expectedEnd.Value Then
                    Dim earlyLeave As TimeSpan = expectedEnd.Value - actualEnd.Value
                    violationText += $"Ранний уход: {earlyLeave.Hours}ч {earlyLeave.Minutes}м. "
                    hasViolation = True
                End If
            End If

            ' Подсвечиваем нарушения
            If hasViolation Then
                HighlightTimeViolations(ws, row, violationText.Trim(), isFriday)
            End If

        Catch ex As Exception
            ' Игнорируем ошибки анализа времени
        End Try
    End Sub

    ' Получает график работы для сотрудника из 15-й колонки
    Private Function GetWorkScheduleForEmployee(ws As Excel.Worksheet, row As Integer) As String
        Dim scheduleText As String = CStr(ExcelUtilities.GetCellValue(ws, row, COL_WORK_SCHEDULE))

        If Not String.IsNullOrEmpty(scheduleText) Then
            Return scheduleText
        End If

        ' Если график работы пустой, возвращаем стандартный
        Return DEFAULT_WORK_SCHEDULE
    End Function

    ' Извлекает время начала работы из текста графика
    Private Function ExtractStartTimeFromSchedule(scheduleText As String) As TimeSpan?
        If String.IsNullOrEmpty(scheduleText) Then Return Nothing

        ' Ищем паттерн времени: "8:00-17:00", "8-00 до 17-00", "с 8:00-17:00"
        Dim timePattern As String = "(\d{1,2})[:-]?(\d{2})\s*(?:до|-|–)\s*(\d{1,2})[:-]?(\d{2})"
        Dim match As Match = Regex.Match(scheduleText, timePattern)

        If match.Success Then
            Dim hour As Integer = Integer.Parse(match.Groups(1).Value)
            Dim minute As Integer = Integer.Parse(match.Groups(2).Value)
            Return New TimeSpan(hour, minute, 0)
        End If

        Return Nothing
    End Function

    ' Извлекает время окончания работы из текста графика
    Private Function ExtractEndTimeFromSchedule(scheduleText As String, isFriday As Boolean) As TimeSpan?
        If String.IsNullOrEmpty(scheduleText) Then Return Nothing

        ' Если пятница, ищем третье время (время окончания в пятницу)
        If isFriday Then
            ' Ищем паттерн для пятницы: "5:00-17:00 (15:45 обед 12:15)" -> 15:45
            ' Или "7:00-16:00(14:45, обед с 12:15)" -> 14:45
            Dim fridayPattern As String = "\((\d{1,2})[:-]?(\d{2})"
            Dim fridayMatch As Match = Regex.Match(scheduleText, fridayPattern)

            If fridayMatch.Success Then
                Dim hour As Integer = Integer.Parse(fridayMatch.Groups(1).Value)
                Dim minute As Integer = Integer.Parse(fridayMatch.Groups(2).Value)
                Return New TimeSpan(hour, minute, 0)
            End If
        End If

        ' Для остальных дней ищем обычное время окончания: "8:00-17:00", "8-00 до 17-00", "с 8:00-17:00"
        Dim timePattern As String = "(\d{1,2})[:-]?(\d{2})\s*(?:до|-|–)\s*(\d{1,2})[:-]?(\d{2})"
        Dim match As Match = Regex.Match(scheduleText, timePattern)

        If match.Success Then
            Dim hour As Integer = Integer.Parse(match.Groups(3).Value)
            Dim minute As Integer = Integer.Parse(match.Groups(4).Value)
            Return New TimeSpan(hour, minute, 0)
        End If

        Return Nothing
    End Function

    ' Парсит время из объекта ячейки Excel (поддерживает числовые значения и текст)
    Private Function ParseTimeFromCellValue(cellValue As Object) As TimeSpan?
        If cellValue Is Nothing Then Return Nothing

        ' Если это число (формат Excel)
        If TypeOf cellValue Is Double Then
            Dim doubleValue As Double = CDbl(cellValue)
            ' Excel хранит время как долю дня (0.5 = 12:00, 0.25 = 6:00)
            If doubleValue >= 0 AndAlso doubleValue <= 1 Then
                Dim totalMinutes As Integer = CInt(doubleValue * 24 * 60)
                Dim hours As Integer = totalMinutes \ 60
                Dim minutes As Integer = totalMinutes Mod 60
                Return New TimeSpan(hours, minutes, 0)
            End If
        End If

        ' Если это строка, пробуем распарсить как текст
        Dim timeValue As String = cellValue.ToString()
        If String.IsNullOrEmpty(timeValue) Then Return Nothing

        ' Сначала пробуем распарсить как число (формат Excel)
        Dim numericValue As Double
        If Double.TryParse(timeValue, numericValue) Then
            ' Excel хранит время как долю дня (0.5 = 12:00, 0.25 = 6:00)
            If numericValue >= 0 AndAlso numericValue <= 1 Then
                Dim totalMinutes As Integer = CInt(numericValue * 24 * 60)
                Dim hours As Integer = totalMinutes \ 60
                Dim minutes As Integer = totalMinutes Mod 60
                Return New TimeSpan(hours, minutes, 0)
            End If
        End If

        ' Если не число, пробуем распарсить как текст: "8:30", "08:30", "8-30"
        Dim timePattern As String = "(\d{1,2})[:-](\d{2})"
        Dim match As Match = Regex.Match(timeValue, timePattern)

        If match.Success Then
            Dim hour As Integer = Integer.Parse(match.Groups(1).Value)
            Dim minute As Integer = Integer.Parse(match.Groups(2).Value)
            Return New TimeSpan(hour, minute, 0)
        End If

        Return Nothing
    End Function

    ' Подсвечивает нарушения времени в ячейках
    Private Sub HighlightTimeViolations(ws As Excel.Worksheet, row As Integer, violationText As String, isFriday As Boolean)
        Try

            ' Разделяем нарушения на опоздания и ранние уходы
            Dim hasLateArrival As Boolean = violationText.Contains("Опоздание")
            Dim hasEarlyLeave As Boolean = violationText.Contains("Ранний уход")

            ' Подсвечиваем опоздания только в колонке "Начало дня"
            If hasLateArrival Then
                Dim startCell As Excel.Range = CType(ws.Cells(row, COL_START_TIME), Excel.Range)
                startCell.Interior.Color = COLOR_LIGHT_RED

                ' Добавляем комментарий с описанием опоздания
                Dim lateComment As String = ExtractLateArrivalText(violationText)
                If Not String.IsNullOrEmpty(lateComment) Then
                    startCell.AddComment(lateComment)
                    
                    ' Настраиваем размер комментария
                    If startCell.Comment IsNot Nothing Then
                        startCell.Comment.Shape.Width = 300
                        startCell.Comment.Shape.Height = 100
                        startCell.Comment.Shape.TextFrame.AutoSize = True
                    End If
                End If

                ExcelUtilities.ReleaseComObject(startCell)
            End If

            ' Подсвечиваем ранние уходы только в колонке "Конец дня"
            If hasEarlyLeave Then
                Dim endCell As Excel.Range = CType(ws.Cells(row, COL_END_TIME), Excel.Range)
                endCell.Interior.Color = COLOR_LIGHT_RED

                ' Добавляем комментарий с описанием раннего ухода
                Dim earlyComment As String = ExtractEarlyLeaveText(violationText)
                If Not String.IsNullOrEmpty(earlyComment) Then
                    ' Если пятница, добавляем информацию о дне недели
                    If isFriday Then
                        earlyComment += " (пятница)"
                    End If
                    endCell.AddComment(earlyComment)
                    
                    ' Настраиваем размер комментария
                    If endCell.Comment IsNot Nothing Then
                        endCell.Comment.Shape.Width = 300
                        endCell.Comment.Shape.Height = 100
                        endCell.Comment.Shape.TextFrame.AutoSize = True
                    End If
                End If

                ExcelUtilities.ReleaseComObject(endCell)
            End If

        Catch ex As Exception
            ' Игнорируем ошибки подсветки
        End Try
    End Sub

    ' Извлекает текст опоздания из общего текста нарушений
    Private Function ExtractLateArrivalText(violationText As String) As String
        If String.IsNullOrEmpty(violationText) Then Return String.Empty

        Dim parts() As String = violationText.Split("."c)
        For Each part In parts
            If part.Trim().StartsWith("Опоздание") Then
                Return part.Trim() & "."
            End If
        Next

        Return String.Empty
    End Function

    ' Извлекает текст раннего ухода из общего текста нарушений
    Private Function ExtractEarlyLeaveText(violationText As String) As String
        If String.IsNullOrEmpty(violationText) Then Return String.Empty

        Dim parts() As String = violationText.Split("."c)
        For Each part In parts
            If part.Trim().StartsWith("Ранний уход") Then
                Return part.Trim() & "."
            End If
        Next

        Return String.Empty
    End Function


End Module
