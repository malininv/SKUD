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

    Private Const ROW_HEADER As Integer = 5
    Private Const ROW_DATA_START As Integer = 6
    Private Const COL_MARKER As Integer = 1 ' A — "ИТОГО"
    Private Const COL_DEPT As Integer = 2   ' B — отдел (первая строка блока)
    Private Const COL_EMPLOYEE As Integer = 3 ' C — ФИО сотрудника
    Private Const COL_DATE As Integer = 6   ' F — дата
    Private Const COL_TIME As Integer = 7   ' G — время
    Private Const COL_START_TIME As Integer = 10 ' J — "Начало дня"
    Private Const COL_END_TIME As Integer = 11   ' K — "Конец дня"
    Private Const COL_OVERTIME As Integer = 13 ' "Фактическая переработка"

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
        Dim calcPrev = app.Calculation
        Dim screenPrev = app.ScreenUpdating
        Dim eventsPrev = app.EnableEvents
        Dim alertsPrev = app.DisplayAlerts
        Dim savedPath As String = Nothing

        Try
            app.Calculation = Excel.XlCalculation.xlCalculationManual
            app.ScreenUpdating = False
            app.EnableEvents = False
            app.DisplayAlerts = False

            Dim wsSource = TryCast(app.ActiveSheet, Excel.Worksheet)
            If wsSource Is Nothing OrElse Not Object.ReferenceEquals(wsSource.Parent, srcWb) Then
                wsSource = CType(srcWb.Sheets(1), Excel.Worksheet)
                wsSource.Activate()
            End If

            ApplyAutoFilter(wsSource, ROW_HEADER)

            Dim lastRow As Integer = wsSource.Cells(wsSource.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            Dim lastCol As Integer = wsSource.Cells(ROW_HEADER, wsSource.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

            ' Create output workbook
            Dim wbNew As Excel.Workbook = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet)
            Dim wsFirst As Excel.Worksheet = CType(wbNew.Sheets(1), Excel.Worksheet)
            wsFirst.Cells.Clear()
            wsFirst.Name = "Отчет"

            Dim baseName As String = Path.GetFileNameWithoutExtension(srcWb.Name)
            Dim newName As String = baseName & "_по_отделам.xlsx"
            Dim saveDir As String = If(String.IsNullOrEmpty(srcWb.Path), app.DefaultFilePath, srcWb.Path)
            Dim savePath As String = Path.Combine(saveDir, newName)
            If File.Exists(savePath) Then File.Delete(savePath)
            wbNew.SaveAs(Filename:=savePath, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook)

            Dim created As New Dictionary(Of String, Excel.Worksheet)(StringComparer.CurrentCulture)
            Dim startCopyRow As Integer = ROW_DATA_START

            For r As Integer = ROW_HEADER + 1 To lastRow
                Dim markerObj As Object = GetCellValue(wsSource, r, COL_MARKER)
                If StringEquals(markerObj, "ИТОГО") Then
                    Dim deptRaw As String = CStr(GetCellValue(wsSource, startCopyRow, COL_DEPT))
                    Dim dept As String = LimitSheetName(deptRaw)
                    Dim timeObj As Object = GetCellValue(wsSource, r, COL_TIME)
                    If IsZeroTime(timeObj) Then dept &= "_нет_прохода"

                    If dept.Length > 0 Then
                        Dim wsTarget As Excel.Worksheet = Nothing
                        If Not created.TryGetValue(dept, wsTarget) OrElse wsTarget Is Nothing Then
                            wsTarget = CreateOrGetSheet(wbNew, dept)
                            ' Header copy
                            Dim headerRowRange As Excel.Range = CType(wsSource.Rows(ROW_HEADER), Excel.Range)
                            Dim destHeaderRow As Excel.Range = CType(wsTarget.Rows(1), Excel.Range)
                            headerRowRange.Copy(Destination:=destHeaderRow)
                            destHeaderRow.Font.Bold = True
                            Marshal.FinalReleaseComObject(headerRowRange)
                            Marshal.FinalReleaseComObject(destHeaderRow)
                            created(dept) = wsTarget
                        End If

                        Dim pasteRow As Integer = wsTarget.Cells(wsTarget.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row + 1
                        Dim srcRange As Excel.Range = wsSource.Range(wsSource.Rows(startCopyRow), wsSource.Rows(r))
                        Dim destPaste As Excel.Range = CType(wsTarget.Rows(pasteRow), Excel.Range)
                        srcRange.Copy(Destination:=destPaste)
                        Marshal.FinalReleaseComObject(srcRange)
                        Marshal.FinalReleaseComObject(destPaste)

                        BeautifySheet(wsTarget)
                        Dim rngFit As Excel.Range = wsTarget.UsedRange
                        rngFit.Columns.AutoFit()
                        Marshal.FinalReleaseComObject(rngFit)
                    End If
                    startCopyRow = r + 1
                End If
            Next

            ' Удаляем ненужные колонки из всех листов перед сохранением
            For Each ws As Excel.Worksheet In wbNew.Sheets
                RemoveUnnecessaryColumns(ws)
            Next

            SortSheetsAlphabetically(wbNew)
            app.StatusBar = $"Готово! Файл сохранён: {newName}"
            srcWb.Activate()
            savedPath = savePath
        Finally
            app.Calculation = calcPrev
            app.ScreenUpdating = screenPrev
            app.EnableEvents = eventsPrev
            app.DisplayAlerts = alertsPrev
        End Try

        Return savedPath
    End Function

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
        Dim calcPrev = app.Calculation
        Dim screenPrev = app.ScreenUpdating
        Dim eventsPrev = app.EnableEvents
        Dim alertsPrev = app.DisplayAlerts

        Dim savedPaths As New List(Of String)()

        Try
            app.Calculation = Excel.XlCalculation.xlCalculationManual
            app.ScreenUpdating = False
            app.EnableEvents = False
            app.DisplayAlerts = False

            Dim wsSource = TryCast(app.ActiveSheet, Excel.Worksheet)
            If wsSource Is Nothing OrElse Not Object.ReferenceEquals(wsSource.Parent, srcWb) Then
                wsSource = CType(srcWb.Sheets(1), Excel.Worksheet)
                wsSource.Activate()
            End If

            ApplyAutoFilter(wsSource, ROW_HEADER)

            Dim lastRow As Integer = wsSource.Cells(wsSource.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            Dim lastCol As Integer = wsSource.Cells(ROW_HEADER, wsSource.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

            Dim saveRoot As String = If(String.IsNullOrEmpty(srcWb.Path), app.DefaultFilePath, srcWb.Path)
            Dim baseNamePerDept As String = Path.GetFileNameWithoutExtension(srcWb.Name)
            Dim targetDir As String = Path.Combine(saveRoot, LimitFileBaseName(baseNamePerDept) & "_по_отделам")
            If Not Directory.Exists(targetDir) Then Directory.CreateDirectory(targetDir)

            Dim deptToWb As New Dictionary(Of String, Excel.Workbook)(StringComparer.CurrentCulture)
            Dim deptToPath As New Dictionary(Of String, String)(StringComparer.CurrentCulture)
            Dim createdSheets As New Dictionary(Of String, Excel.Worksheet)(StringComparer.CurrentCulture)
            Dim deptHasAnySheet As New HashSet(Of String)(StringComparer.CurrentCulture)

            Dim startCopyRow As Integer = ROW_DATA_START
            For r As Integer = ROW_HEADER + 1 To lastRow
                Dim markerObj As Object = GetCellValue(wsSource, r, COL_MARKER)
                If StringEquals(markerObj, "ИТОГО") Then
                    Dim deptRaw As String = CStr(GetCellValue(wsSource, startCopyRow, COL_DEPT))
                    Dim deptBase As String = LimitSheetName(deptRaw)
                    If deptBase.Length = 0 Then
                        startCopyRow = r + 1
                        Continue For
                    End If

                    Dim sheetName As String = deptBase
                    Dim timeObj As Object = GetCellValue(wsSource, r, COL_TIME)
                    If IsZeroTime(timeObj) Then sheetName &= "_нет_прохода"

                    ' Get/create workbook per dept
                    Dim wbDept As Excel.Workbook = Nothing
                    Dim wbPath As String = Nothing
                    If Not deptToWb.TryGetValue(deptBase, wbDept) Then
                        wbDept = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet)
                        CType(wbDept.Sheets(1), Excel.Worksheet).Name = "Отчет"
                        wbPath = Path.Combine(targetDir, LimitFileBaseName(deptBase) & ".xlsx")
                        If File.Exists(wbPath) Then File.Delete(wbPath)
                        wbDept.SaveAs(Filename:=wbPath, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook)
                        deptToWb(deptBase) = wbDept
                        deptToPath(deptBase) = wbPath
                        savedPaths.Add(wbPath)
                    Else
                        wbPath = deptToPath(deptBase)
                    End If

                    ' Get/create sheet inside dept workbook
                    Dim key As String = deptBase & "|" & sheetName
                    Dim wsTarget As Excel.Worksheet = Nothing
                    If createdSheets.TryGetValue(key, wsTarget) Then
                        ' already exists
                    Else
                        If Not deptHasAnySheet.Contains(deptBase) AndAlso wbDept.Sheets.Count = 1 AndAlso CType(wbDept.Sheets(1), Excel.Worksheet).Name = "Отчет" Then
                            wsTarget = CType(wbDept.Sheets(1), Excel.Worksheet)
                            wsTarget.Name = sheetName
                        Else
                            wsTarget = CType(wbDept.Sheets.Add(After:=wbDept.Sheets(wbDept.Sheets.Count)), Excel.Worksheet)
                            wsTarget.Name = sheetName
                        End If
                        Dim headerRowRange As Excel.Range = CType(wsSource.Rows(ROW_HEADER), Excel.Range)
                        Dim destHeaderRow As Excel.Range = CType(wsTarget.Rows(1), Excel.Range)
                        headerRowRange.Copy(Destination:=destHeaderRow)
                        destHeaderRow.Font.Bold = True
                        Marshal.FinalReleaseComObject(headerRowRange)
                        Marshal.FinalReleaseComObject(destHeaderRow)
                        createdSheets(key) = wsTarget
                        deptHasAnySheet.Add(deptBase)
                    End If

                    ' Insert block
                    Dim pasteRow As Integer = wsTarget.Cells(wsTarget.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row + 1
                    Dim srcRange As Excel.Range = wsSource.Range(wsSource.Rows(startCopyRow), wsSource.Rows(r))
                    Dim destPaste As Excel.Range = CType(wsTarget.Rows(pasteRow), Excel.Range)
                    srcRange.Copy(Destination:=destPaste)
                    Marshal.FinalReleaseComObject(srcRange)
                    Marshal.FinalReleaseComObject(destPaste)

                    BeautifySheet(wsTarget)
                    Dim rngFit As Excel.Range = wsTarget.UsedRange
                    rngFit.Columns.AutoFit()
                    Marshal.FinalReleaseComObject(rngFit)

                    ' Save dept workbook incrementally
                    wbDept.Save()

                    startCopyRow = r + 1
                End If
            Next

            ' Finalize: sort sheets and close dept workbooks
            For Each kv In deptToWb
                Dim wbDept As Excel.Workbook = kv.Value

                ' Удаляем ненужные колонки из всех листов перед сохранением
                For Each ws As Excel.Worksheet In wbDept.Sheets
                    RemoveUnnecessaryColumns(ws)
                Next

                SortSheetsAlphabetically(wbDept)
                wbDept.Save()
                wbDept.Close(SaveChanges:=False)
                Marshal.FinalReleaseComObject(wbDept)
            Next

            app.StatusBar = $"Готово! Создано файлов: {savedPaths.Count}. Папка: {targetDir}"
            srcWb.Activate()
        Finally
            app.Calculation = calcPrev
            app.ScreenUpdating = screenPrev
            app.EnableEvents = eventsPrev
            app.DisplayAlerts = alertsPrev
        End Try

        Return savedPaths
    End Function

    ' ====================== Helpers ======================
    Private Sub ApplyAutoFilter(ws As Excel.Worksheet, headerRow As Integer)
        ' Safe/optional: activates sheet, tries main range then UsedRange; ignores errors.
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
                    If ur IsNot Nothing Then Marshal.FinalReleaseComObject(ur)
                End Try
            Finally
                If rng IsNot Nothing Then Marshal.FinalReleaseComObject(rng)
            End Try
        Catch
            ' ignore
        End Try
    End Sub

    Private Sub BeautifySheet(ws As Excel.Worksheet)
        Dim lastRow As Integer = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        Dim paleYellow As Integer = ColorTranslator.ToOle(Color.FromArgb(255, 255, 204)) ' #FFFFCC
        Dim red As Integer = ColorTranslator.ToOle(Color.Red)

        ' Сначала удаляем выходные строки
        For r As Integer = lastRow To 2 Step -1
            Dim dateVal As Object = GetCellValue(ws, r, COL_DATE)
            Dim timeVal As Object = GetCellValue(ws, r, COL_TIME)
            'удаляем выходные
            If IsWeekend(dateVal) AndAlso IsZeroTime(timeVal) Then
                CType(ws.Rows(r), Excel.Range).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                Continue For
            End If
        Next

        ' Теперь обрабатываем оставшиеся строки
        lastRow = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row

        ' ПЕРВЫЙ ЭТАП: Обработка всех строк (окрашивание, анализ времени, установка 0)
        For r As Integer = lastRow To 2 Step -1
            Dim dateVal As Object = GetCellValue(ws, r, COL_DATE)
            Dim timeVal As Object = GetCellValue(ws, r, COL_TIME)

            'красим 0 проходы
            If IsZeroTime(timeVal) Then
                Dim tcell As Excel.Range = CType(ws.Cells(r, COL_TIME), Excel.Range)
                tcell.Interior.Color = paleYellow ' фон желтый
                tcell.Font.Color = ColorTranslator.ToOle(Color.Red) ' шрифт красный
                Marshal.FinalReleaseComObject(tcell)
            End If

            ' ==================== АНАЛИЗ ОПОЗДАНИЙ И РАННИХ УХОДОВ ====================
            ' Анализируем время прихода/ухода только для рабочих дней и не для строк ИТОГО
            If Not IsWeekend(dateVal) AndAlso Not StringEquals(ws.Cells(r, COL_MARKER).Value2, "ИТОГО") Then
                AnalyzeWorkTimeViolations(ws, r)

                ' Проверяем, есть ли вход и выход для сотрудника
                Dim startTime As String = CStr(GetCellValue(ws, r, COL_START_TIME))
                Dim endTime As String = CStr(GetCellValue(ws, r, COL_END_TIME))

                ' Если нет входа или выхода, ставим 0 в фактическую переработку
                If String.IsNullOrEmpty(startTime) OrElse String.IsNullOrEmpty(endTime) OrElse
                   startTime.Contains("Нет входа") OrElse endTime.Contains("Нет выход") Then
                    Dim overtimeCell As Excel.Range = CType(ws.Cells(r, COL_OVERTIME), Excel.Range)
                    overtimeCell.Value2 = 0
                    Marshal.FinalReleaseComObject(overtimeCell)

                    ' Отладка: логируем установку 0 (без модального окна)
                    ' MessageBox.Show($"Строка {r}: установлен 0 для 'Нет входа/выхода'. StartTime: '{startTime}', EndTime: '{endTime}'", "Отладка", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Next

        ' ВТОРОЙ ЭТАП: Суммирование для строк ИТОГО
        For r As Integer = lastRow To 2 Step -1

            ' --- Обработка "Фактическая переработка" ТОЛЬКО для строки ИТОГО (13-й столбец) ---
            Dim marker As Object = ws.Cells(r, COL_MARKER).Value2
            If Not IsNothing(marker) AndAlso String.Equals(CStr(marker), "ИТОГО", StringComparison.CurrentCultureIgnoreCase) Then

                ' Добавляем формулу для пересчета фактической переработки по сотруднику
                Dim ocell As Excel.Range = CType(ws.Cells(r, COL_OVERTIME), Excel.Range) ' 13-й столбец

                ' Ищем начало блока сотрудника (предыдущая строка ИТОГО + 1)
                Dim formulaStartRow As Integer = r - 1
                While formulaStartRow > 1 AndAlso Not StringEquals(ws.Cells(formulaStartRow, COL_MARKER).Value2, "ИТОГО")
                    formulaStartRow -= 1
                End While

                ' Если нашли предыдущую строку ИТОГО, начинаем с следующей строки
                If StringEquals(ws.Cells(formulaStartRow, COL_MARKER).Value2, "ИТОГО") Then
                    formulaStartRow += 1
                End If

                ' Исправляем логику: если не нашли предыдущий ИТОГО, начинаем с первой строки данных
                If formulaStartRow <= 1 Then
                    formulaStartRow = 2 ' Первая строка данных после заголовка
                End If

                ' Информация о диапазоне для комментария
                Dim rangeInfo As String = $"ИТОГО строка: {r}, Начало блока: {formulaStartRow}, Конец: {r - 1}"

                ' Проверяем, что есть строки для суммирования
                If formulaStartRow < r - 1 Then
                    ' Считаем сумму программно
                    Dim totalHours As Double = 0
                    Dim debugValues As New List(Of String)

                    For row As Integer = formulaStartRow To r - 1
                        Dim cellValue As Object = GetCellValue(ws, row, COL_OVERTIME)
                        If cellValue IsNot Nothing Then
                            ' Пробуем преобразовать в число разными способами
                            Dim numericValue As Double = 0
                            If IsNumeric(cellValue) Then
                                numericValue = CDbl(cellValue)
                                debugValues.Add($"Строка {row}: {cellValue} (число) -> {numericValue}")
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
                                            debugValues.Add($"Строка {row}: {cellValue} (текст) -> {numericValue}")
                                        End If
                                    Catch
                                        debugValues.Add($"Строка {row}: {cellValue} (ошибка парсинга)")
                                    End Try
                                Else
                                    debugValues.Add($"Строка {row}: {cellValue} (не время)")
                                End If
                            End If
                            totalHours += numericValue
                        Else
                            debugValues.Add($"Строка {row}: пустое значение")
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

                    ' Добавляем комментарий с информацией о диапазоне
                    Try
                        Dim commentText As String = $"Диапазон суммирования: строки {formulaStartRow} - {r - 1}" & Environment.NewLine & rangeInfo & Environment.NewLine & "Итого: " & totalHours.ToString("F6")
                        If debugValues.Count > 0 Then
                            commentText &= Environment.NewLine & "Значения:" & Environment.NewLine & String.Join(Environment.NewLine, debugValues)
                        End If
                        ' Удаляем существующий комментарий, если есть
                        If ocell.Comment IsNot Nothing Then
                            ocell.Comment.Delete()
                        End If
                        ' Добавляем новый комментарий
                        ocell.AddComment(commentText)

                        ' Расширяем размер комментария для лучшей видимости
                        If ocell.Comment IsNot Nothing Then
                            ocell.Comment.Shape.Width = 400
                            ocell.Comment.Shape.Height = 300
                            ocell.Comment.Shape.TextFrame.AutoSize = True
                        End If
                    Catch ex As Exception
                        ' Игнорируем ошибки с комментариями
                    End Try
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
                        ocell.Interior.Color = ColorTranslator.ToOle(Color.LimeGreen)        ' > 0 → зелёный
                    ElseIf hours >= -1 Then
                        ocell.Interior.Color = ColorTranslator.ToOle(Color.Yellow)           ' -1..0 → жёлтый (напр. -0,27)
                    Else
                        ocell.Interior.Color = ColorTranslator.ToOle(Color.Red)              ' < -1 → красный (напр. -1,27 или -167)
                    End If
                End If

                Marshal.FinalReleaseComObject(ocell)
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

            Runtime.InteropServices.Marshal.FinalReleaseComObject(borders)
            Runtime.InteropServices.Marshal.FinalReleaseComObject(rngAll)
        End If

    End Sub

    ' Удаляет ненужные колонки из листа и переносит "ИТОГО" во вторую колонку
    Private Sub RemoveUnnecessaryColumns(ws As Excel.Worksheet)
        Try
            ' Переименовываем колонку "Мягкие прогулы" в "Находился вне здания" и добавляем комментарий (9-я колонка)
            Try
                Dim headerRow As Integer = 1
                Dim headerCell As Excel.Range = CType(ws.Cells(headerRow, 9), Excel.Range)
                Dim headerValue As Object = GetCellValue(ws, headerRow, 9)

                If Not IsNothing(headerValue) AndAlso CStr(headerValue).Contains("Мягких прогулов") Then
                    ' Переименовываем заголовок
                    headerCell.Value2 = "Находился вне здания"

                    ' Добавляем комментарий
                    Dim commentText As String = "Время, которое сотрудник находился вне здания в необеденное время"
                    If headerCell.Comment IsNot Nothing Then
                        headerCell.Comment.Delete()
                    End If
                    headerCell.AddComment(commentText)

                    ' Расширяем размер комментария и выравниваем ширину столбца
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
                    Dim columnRange As Excel.Range = CType(ws.Columns(9), Excel.Range)
                    columnRange.AutoFit()
                    Marshal.FinalReleaseComObject(columnRange)
                End If

                Marshal.FinalReleaseComObject(headerCell)
            Catch ex As Exception
                ' Игнорируем ошибки переименования
            End Try

            ' Сначала переносим "ИТОГО" из первой колонки во вторую
            Dim lastRow As Integer = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            For r As Integer = ROW_DATA_START To lastRow
                Dim markerObj As Object = GetCellValue(ws, r, 1) ' Первая колонка
                If Not IsNothing(markerObj) AndAlso String.Equals(CStr(markerObj), "ИТОГО", StringComparison.CurrentCultureIgnoreCase) Then
                    ' Копируем "ИТОГО" во вторую колонку
                    Dim targetCell As Excel.Range = CType(ws.Cells(r, 2), Excel.Range)
                    targetCell.Value2 = "ИТОГО"
                    Marshal.FinalReleaseComObject(targetCell)

                    ' Очищаем первую колонку
                    Dim sourceCell As Excel.Range = CType(ws.Cells(r, 1), Excel.Range)
                    sourceCell.Value2 = ""
                    Marshal.FinalReleaseComObject(sourceCell)
                End If
            Next

            ' Удаляем колонки в обратном порядке, чтобы не сбить нумерацию
            ' Удаляем "Работа в праздничные дни" (12-я колонка)
            Dim holidayCol As Excel.Range = CType(ws.Columns(12), Excel.Range)
            holidayCol.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
            Marshal.FinalReleaseComObject(holidayCol)

            ' Удаляем "Прогулял" (8-я колонка)
            Dim absentCol As Excel.Range = CType(ws.Columns(8), Excel.Range)
            absentCol.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
            Marshal.FinalReleaseComObject(absentCol)

            ' Удаляем "Таб #" (5-я колонка)
            Dim tabCol As Excel.Range = CType(ws.Columns(5), Excel.Range)
            tabCol.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
            Marshal.FinalReleaseComObject(tabCol)

            ' Удаляем первую колонку (теперь пустую)
            Dim firstCol As Excel.Range = CType(ws.Columns(1), Excel.Range)
            firstCol.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
            Marshal.FinalReleaseComObject(firstCol)


        Catch ex As Exception
            ' Игнорируем ошибки удаления колонок
        End Try
    End Sub

    Private Function CreateOrGetSheet(wb As Excel.Workbook, baseName As String) As Excel.Worksheet
        ' Create sheet; if name collision — add (2), (3)...
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

    Private Function TryGetSheet(wb As Excel.Workbook, name As String) As Excel.Worksheet
        For Each sh As Object In wb.Sheets
            Dim ws = TryCast(sh, Excel.Worksheet)
            If ws IsNot Nothing AndAlso String.Equals(ws.Name, name, StringComparison.CurrentCulture) Then
                Return ws
            End If
        Next
        Return Nothing
    End Function

    Private Function MakeUniqueSheetName(baseName As String, index As Integer) As String
        Dim core As String = baseName
        Dim suffix As String = " (" & index.ToString(CultureInfo.CurrentCulture) & ")"
        Dim maxLen As Integer = 31
        If core.Length + suffix.Length > maxLen Then
            core = core.Substring(0, Math.Max(0, maxLen - suffix.Length))
        End If
        Return core & suffix
    End Function

    Private Sub SortSheetsAlphabetically(wb As Excel.Workbook)
        Dim list As New List(Of Excel.Worksheet)
        For Each sh As Object In wb.Sheets
            Dim ws = TryCast(sh, Excel.Worksheet)
            If ws IsNot Nothing Then list.Add(ws)
        Next
        list.Sort(Function(a, b) String.Compare(a.Name, b.Name, StringComparison.CurrentCulture))
        For i As Integer = 0 To list.Count - 1
            Dim ws As Excel.Worksheet = list(i)
            ws.Move(After:=wb.Sheets(i + 1))
        Next
    End Sub

    Private Function GetCellValue(ws As Excel.Worksheet, row As Integer, col As Integer) As Object
        Dim rng As Excel.Range = CType(ws.Cells(row, col), Excel.Range)
        Dim v As Object = rng.Value2
        Marshal.FinalReleaseComObject(rng)
        Return v
    End Function

    Private Function StringEquals(v As Object, expected As String) As Boolean
        If v Is Nothing Then Return False
        Dim s As String = CStr(v)
        Return String.Compare(s, expected, True, CultureInfo.CurrentCulture) = 0
    End Function

    Private Function IsZeroTime(v As Object) As Boolean
        If v Is Nothing Then Return False
        If TypeOf v Is Double Then
            Return Math.Abs(CDbl(v)) < 0.0000001R
        End If
        Dim s As String = CStr(v).Trim()
        Return s.StartsWith("0:00", StringComparison.CurrentCulture)
    End Function

    Private Function IsWeekend(v As Object) As Boolean
        If v Is Nothing Then Return False
        Dim dt As Date
        If TypeOf v Is Double Then
            dt = Date.FromOADate(CDbl(v))
        ElseIf Not Date.TryParse(CStr(v), dt) Then
            Return False
        End If
        Dim mondayBased As Integer = ((CInt(dt.DayOfWeek) + 6) Mod 7) + 1 ' 1=Mon .. 7=Sun
        Return (mondayBased = 6 OrElse mondayBased = 7)
    End Function

    Private Function LimitSheetName(proposed As String) As String
        Dim name As String = If(proposed, String.Empty).Trim()
        ' Excel допускает до 31 символа в имени листа
        If name.Length > 18 Then name = name.Substring(0, 18)
        Dim banned As String = "\/?*[]:"
        For Each ch As Char In banned
            name = name.Replace(ch.ToString(), "_")
        Next
        While name.EndsWith(" ") OrElse name.EndsWith(".")
            name = If(name.Length > 1, name.Substring(0, name.Length - 1), String.Empty)
            If name.Length = 0 Then Exit While
        End While
        If String.IsNullOrWhiteSpace(name) Then name = "Лист"
        Return name
    End Function

    Private Function LimitFileBaseName(proposed As String) As String
        Dim name As String = If(proposed, String.Empty).Trim()
        Dim maxLen As Integer = 50
        If name.Length > maxLen Then name = name.Substring(0, maxLen)
        Dim banned As String = "\/:*?""<>|" ' note doubled quote
        For Each ch As Char In banned
            name = name.Replace(ch.ToString(), "_")
        Next
        While name.EndsWith(" ") OrElse name.EndsWith(".")
            name = If(name.Length > 1, name.Substring(0, name.Length - 1), String.Empty)
            If name.Length = 0 Then Exit While
        End While
        If String.IsNullOrWhiteSpace(name) Then name = "Отчет"
        Return name
    End Function

    ' ==================== АНАЛИЗ ОПОЗДАНИЙ И РАННИХ УХОДОВ ====================
    ' Анализирует время прихода/ухода и подсвечивает нарушения
    Private Sub AnalyzeWorkTimeViolations(ws As Excel.Worksheet, row As Integer)
        Try
            ' Читаем ФИО сотрудника
            Dim fio As String = CStr(GetCellValue(ws, row, COL_EMPLOYEE))
            If String.IsNullOrEmpty(fio) Then Return

            ' Читаем время прихода и ухода
            Dim startTime As String = CStr(GetCellValue(ws, row, COL_START_TIME))
            Dim endTime As String = CStr(GetCellValue(ws, row, COL_END_TIME))

            ' Пропускаем если нет данных о времени
            If String.IsNullOrEmpty(startTime) AndAlso String.IsNullOrEmpty(endTime) Then Return
            If startTime.Contains("Нет входа") AndAlso endTime.Contains("Нет выход") Then Return

            ' Получаем график работы для сотрудника из колонки "График работы"
            Dim workSchedule As String = GetWorkScheduleForEmployee(ws, row)
            If String.IsNullOrEmpty(workSchedule) Then Return

            ' Если график стандартный (не проставлен), не анализируем
            If workSchedule = "рабочий график: 8:00-17:00" Then Return

            ' Получаем дату для определения дня недели
            Dim dayValue As Date? = LeaveReasonFiller.ReadDate(ws, row, COL_DATE)
            Dim isFriday As Boolean = dayValue.HasValue AndAlso dayValue.Value.DayOfWeek = DayOfWeek.Friday

            ' Извлекаем время начала и окончания работы из графика
            Dim expectedStart As TimeSpan? = ExtractStartTimeFromSchedule(workSchedule)
            Dim expectedEnd As TimeSpan? = ExtractEndTimeFromSchedule(workSchedule, isFriday)

            ' Получаем объекты ячеек для анализа
            Dim startTimeObj As Object = GetCellValue(ws, row, COL_START_TIME)
            Dim endTimeObj As Object = GetCellValue(ws, row, COL_END_TIME)


            Dim hasViolation As Boolean = False
            Dim violationText As String = ""

            ' Анализ времени прихода
            If Not String.IsNullOrEmpty(startTime) AndAlso Not startTime.Contains("Нет входа") AndAlso expectedStart.HasValue Then
                Dim actualStart As TimeSpan? = ParseTimeFromCellValue(startTimeObj)

                If actualStart.HasValue AndAlso actualStart.Value > expectedStart.Value Then
                    Dim delay As TimeSpan = actualStart.Value - expectedStart.Value
                    violationText += $"Опоздание: {delay.Hours}ч {delay.Minutes}м. "
                    hasViolation = True
                End If
            End If

            ' Анализ времени ухода
            If Not String.IsNullOrEmpty(endTime) AndAlso Not endTime.Contains("Нет выход") AndAlso expectedEnd.HasValue Then
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
        ' Всегда читаем из 15-й колонки (колонка O)
        Const SCHEDULE_COLUMN As Integer = 15

        Dim scheduleText As String = CStr(GetCellValue(ws, row, SCHEDULE_COLUMN))

        If Not String.IsNullOrEmpty(scheduleText) Then
            Return scheduleText
        End If

        ' Если график работы пустой, возвращаем стандартный
        Return "рабочий график: 8:00-17:00"
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
            ' Устанавливаем светло-красный фон для нарушений
            Dim lightRed As Integer = ColorTranslator.ToOle(Color.FromArgb(255, 200, 200))

            ' Разделяем нарушения на опоздания и ранние уходы
            Dim hasLateArrival As Boolean = violationText.Contains("Опоздание")
            Dim hasEarlyLeave As Boolean = violationText.Contains("Ранний уход")

            ' Подсвечиваем опоздания только в колонке "Начало дня"
            If hasLateArrival Then
                Dim startCell As Excel.Range = CType(ws.Cells(row, COL_START_TIME), Excel.Range)
                startCell.Interior.Color = lightRed

                ' Добавляем комментарий с описанием опоздания
                Dim lateComment As String = ExtractLateArrivalText(violationText)
                If Not String.IsNullOrEmpty(lateComment) Then
                    startCell.AddComment(lateComment)
                End If

                Marshal.FinalReleaseComObject(startCell)
            End If

            ' Подсвечиваем ранние уходы только в колонке "Конец дня"
            If hasEarlyLeave Then
                Dim endCell As Excel.Range = CType(ws.Cells(row, COL_END_TIME), Excel.Range)
                endCell.Interior.Color = lightRed

                ' Добавляем комментарий с описанием раннего ухода
                Dim earlyComment As String = ExtractEarlyLeaveText(violationText)
                If Not String.IsNullOrEmpty(earlyComment) Then
                    ' Если пятница, добавляем информацию о дне недели
                    If isFriday Then
                        earlyComment += " (пятница)"
                    End If
                    endCell.AddComment(earlyComment)
                End If

                Marshal.FinalReleaseComObject(endCell)
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
