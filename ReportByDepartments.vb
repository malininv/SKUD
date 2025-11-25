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
    Private Const COL_START_TIME As Integer = 11 ' J — "Начало дня"
    Private Const COL_END_TIME As Integer = 12   ' K — "Конец дня"
    Private Const COL_OVERTIME As Integer = 14 ' "Фактическая переработка"

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

                            ' Переименовываем заголовки
                            RenameHeaders(wsTarget)

                            created(dept) = wsTarget
                        End If

                        Dim pasteRow As Integer = wsTarget.Cells(wsTarget.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row + 1
                        Dim srcRange As Excel.Range = wsSource.Range(wsSource.Rows(startCopyRow), wsSource.Rows(r))
                        Dim destPaste As Excel.Range = CType(wsTarget.Rows(pasteRow), Excel.Range)
                        srcRange.Copy(Destination:=destPaste)
                        Marshal.FinalReleaseComObject(srcRange)
                        Marshal.FinalReleaseComObject(destPaste)

                        ' BeautifySheet будет вызван после цикла для всех листов
                    End If
                    startCopyRow = r + 1
                End If
            Next

            ' Обрабатываем все созданные листы (BeautifySheet вызывается только один раз для каждого листа)
            For Each kvp In created
                Dim wsTarget As Excel.Worksheet = kvp.Value
                If wsTarget IsNot Nothing Then
                    BeautifySheet(wsTarget)

                    Dim rngFit As Excel.Range = wsTarget.UsedRange
                    rngFit.Columns.AutoFit()
                    Marshal.FinalReleaseComObject(rngFit)

                    ' Закрепляем заголовок
                    FreezeHeaderRow(wsTarget)
                End If
            Next

            ' Удаляем ненужные колонки из всех листов перед сохранением
            For Each ws As Excel.Worksheet In wbNew.Sheets
                RemoveUnnecessaryColumns(ws)
            Next

            ' Удаляем пустые листы "Отчет"
            RemoveEmptyReportSheets(wbNew)

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

                        ' Переименовываем заголовки
                        RenameHeaders(wsTarget)

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

                    ' BeautifySheet будет вызван после цикла для всех листов

                    startCopyRow = r + 1
                End If
            Next

            ' Finalize: sort sheets and close dept workbooks
            For Each kv In deptToWb
                Dim wbDept As Excel.Workbook = kv.Value

                ' Обрабатываем все листы (BeautifySheet вызывается только один раз для каждого листа)
                For Each ws As Excel.Worksheet In wbDept.Sheets
                    If ws.Name <> "Отчет" OrElse ws.UsedRange.Rows.Count > 1 Then
                        BeautifySheet(ws)

                        Dim rngFit As Excel.Range = ws.UsedRange
                        rngFit.Columns.AutoFit()
                        Marshal.FinalReleaseComObject(rngFit)

                        ' Закрепляем заголовок
                        FreezeHeaderRow(ws)
                    End If
                Next

                ' Удаляем ненужные колонки из всех листов
                For Each ws As Excel.Worksheet In wbDept.Sheets
                    RemoveUnnecessaryColumns(ws)
                Next

                ' Удаляем пустые листы "Отчет"
                RemoveEmptyReportSheets(wbDept)

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

    Private Sub FreezeHeaderRow(ws As Excel.Worksheet)
        ' Закрепляем первую строку (заголовок) при пролистывании
        Try
            ws.Activate()
            ws.Range("A2").Select()
            ws.Application.ActiveWindow.FreezePanes = True
        Catch ex As Exception
            ' Игнорируем ошибки закрепления панелей
        End Try
    End Sub

    Private Sub RenameHeaders(ws As Excel.Worksheet)
        ' Единая функция для переименования всех заголовков
        Try
            Dim lastCol As Integer = ws.Cells(1, ws.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

            For col As Integer = 1 To lastCol
                Dim headerCell As Excel.Range = CType(ws.Cells(1, col), Excel.Range)
                Dim headerValue As Object = GetCellValue(ws, 1, col)

                If Not IsNothing(headerValue) Then
                    Dim headerText As String = CStr(headerValue).Trim()
                    Dim newHeaderText As String = headerText

                    ' Переименовываем заголовки
                    If headerText.Contains("Прогулял") Then
                        newHeaderText = "Находился вне здания (Ч:М)"

                        ' Добавляем комментарий для этого заголовка
                        If headerCell.Comment IsNot Nothing Then
                            headerCell.Comment.Delete()
                        End If
                        headerCell.AddComment("Время, которое сотрудник находился вне здания, исключая обеденное время")

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

                    ElseIf headerText.Contains("Находился в здании") AndAlso Not headerText.Contains("(Ч:М)") Then
                        newHeaderText = headerText & " (Ч:М)"

                        ' Добавляем комментарий для этого заголовка
                        If headerCell.Comment IsNot Nothing Then
                            headerCell.Comment.Delete()
                        End If
                        headerCell.AddComment("Время, которое сотрудник находился в здании, исключая обеденное время + переработка")

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
                    ElseIf headerText.Contains("Фактическая переработка") AndAlso Not headerText.Contains("(Ч:М)") Then
                        newHeaderText = headerText & " (Ч:М)"
                        ' Добавляем комментарий для этого заголовка
                        If headerCell.Comment IsNot Nothing Then
                            headerCell.Comment.Delete()
                        End If
                        headerCell.AddComment("(Утренняя переработка + Вечерняя переработка) - Находился вне здания (исключая время обеда)")

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
                    End If

                    ' Обновляем заголовок, если он изменился
                    If newHeaderText <> headerText Then
                        headerCell.Value2 = newHeaderText
                    End If
                End If

                Marshal.FinalReleaseComObject(headerCell)
            Next
        Catch ex As Exception
            ' Игнорируем ошибки переименования заголовков
        End Try
    End Sub


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

        ' ЭТАП 1: Удаление выходных строк (только если сотрудник не работал)
        lastRow = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        For r As Integer = lastRow To 2 Step -1
            Dim dateVal As Object = GetCellValue(ws, r, COL_DATE)
            Dim timeVal As Object = GetCellValue(ws, r, COL_TIME)
            Dim startTimeVal As Object = GetCellValue(ws, r, COL_START_TIME)
            Dim endTimeVal As Object = GetCellValue(ws, r, COL_END_TIME)

            ' Проверяем, является ли строка строкой ИТОГО
            Dim markerVal As Object = GetCellValue(ws, r, COL_MARKER)
            Dim isTotalRow As Boolean = StringEquals(markerVal, "ИТОГО")

            ' Удаляем выходные только если:
            ' 1. Это НЕ строка ИТОГО
            ' 2. День является выходным
            ' 3. Время в здании равно нулю
            ' 4. Нет данных о начале/конце рабочего дня (сотрудник не работал)
            If Not isTotalRow AndAlso IsWeekend(dateVal) AndAlso IsZeroTime(timeVal) Then
                Dim startTimeStr As String = If(startTimeVal Is Nothing, "", CStr(startTimeVal).Trim())
                Dim endTimeStr As String = If(endTimeVal Is Nothing, "", CStr(endTimeVal).Trim())

                ' Удаляем только если нет записей о работе
                If String.IsNullOrEmpty(startTimeStr) OrElse
                   startTimeStr.Contains("Нет входа") OrElse
                   String.IsNullOrEmpty(endTimeStr) OrElse
                   endTimeStr.Contains("Нет выход") Then
                    CType(ws.Rows(r), Excel.Range).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                    Continue For
                End If
            End If
        Next

        ' ЭТАП 2: Добавление времени обеда ко всем строкам переработки
        lastRow = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        
        For r As Integer = 2 To lastRow
            ' Пропускаем строки ИТОГО
            Dim markerVal As Object = GetCellValue(ws, r, COL_MARKER)
            If StringEquals(markerVal, "ИТОГО") Then Continue For

            ' Проверяем, выходной ли день
            Dim dateVal As Object = GetCellValue(ws, r, COL_DATE)
            Dim isWeekendDay As Boolean = IsWeekend(dateVal)

            ' Проверяем, есть ли причина отсутствия
            Dim hasAbsenceReason As Boolean = False
            Dim lastCol As Integer = ws.Cells(r, ws.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
            For col As Integer = 1 To lastCol
                Dim headerValue As Object = GetCellValue(ws, 1, col)
                If headerValue IsNot Nothing AndAlso CStr(headerValue).Trim().Contains("Причина отсутствия") Then
                    Dim reasonValue As Object = GetCellValue(ws, r, col)
                    If reasonValue IsNot Nothing AndAlso Not String.IsNullOrEmpty(CStr(reasonValue).Trim()) Then
                        hasAbsenceReason = True
                    End If
                    Exit For
                End If
            Next

            ' Для выходных и дней с причиной отсутствия не добавляем обед
            If isWeekendDay OrElse hasAbsenceReason Then
                Continue For
            End If

            ' Читаем текущее значение переработки
            Dim overtimeCell As Excel.Range = CType(ws.Cells(r, COL_OVERTIME), Excel.Range)
            Dim oldHours As Double = ParseOvertimeHours(overtimeCell)

            ' Получаем максимальное время обеда из графика работы
            Dim workSchedule As String = GetWorkScheduleForEmployee(ws, r)
            Dim maxLunchMinutes As Integer = ExtractLunchMinutesFromSchedule(workSchedule)

            ' Находим колонку "Прогулял" / "Находился вне здания" и читаем время обеда
            Dim lunchHours As Double = 0
            For col As Integer = 1 To lastCol
                Dim headerValue As Object = GetCellValue(ws, 1, col)
                If headerValue IsNot Nothing Then
                    Dim headerText As String = CStr(headerValue).Trim()
                    If headerText.Contains("Находился вне здания") OrElse headerText.Contains("Прогулял") Then
                        Dim lunchCell As Excel.Range = CType(ws.Cells(r, col), Excel.Range)
                        Dim lunchValue As Double = ParseOvertimeHours(lunchCell)
                        Marshal.FinalReleaseComObject(lunchCell)

                        ' Ограничиваем обед максимальным значением из графика
                        If lunchValue > 0 Then
                            lunchHours = Math.Min(lunchValue, maxLunchMinutes / 60.0)
                        End If
                        Exit For
                    End If
                End If
            Next

            ' Добавляем обед к переработке
            If lunchHours > 0 Then
                Dim newHours As Double = oldHours + lunchHours
                
                ' Записываем новое значение в формате Ч:ММ
                Dim totalMinutes As Integer = CInt(Math.Abs(newHours) * 60)
                Dim resultHours As Integer = totalMinutes \ 60
                Dim resultMinutes As Integer = totalMinutes Mod 60

                Dim timeString As String
                If newHours < 0 Then
                    timeString = $"-{resultHours}:{resultMinutes:D2}"
                Else
                    timeString = $"{resultHours}:{resultMinutes:D2}"
                End If

                overtimeCell.NumberFormat = "@"
                overtimeCell.Value = timeString
                overtimeCell.HorizontalAlignment = If(newHours >= 0, Excel.XlHAlign.xlHAlignRight, Excel.XlHAlign.xlHAlignLeft)
            End If

            Marshal.FinalReleaseComObject(overtimeCell)
        Next

        ' ЭТАП 3: Корректировка "Находился в здании" с учетом обеда
        lastRow = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        
        For r As Integer = 2 To lastRow
            ' Пропускаем строки ИТОГО
            Dim markerVal As Object = GetCellValue(ws, r, COL_MARKER)
            If StringEquals(markerVal, "ИТОГО") Then Continue For

            ' Проверяем, выходной ли день
            Dim dateVal As Object = GetCellValue(ws, r, COL_DATE)
            Dim isWeekendDay As Boolean = IsWeekend(dateVal)

            ' Проверяем, есть ли причина отсутствия
            Dim hasAbsenceReason As Boolean = False
            Dim lastCol As Integer = ws.Cells(r, ws.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
            For col As Integer = 1 To lastCol
                Dim headerValue As Object = GetCellValue(ws, 1, col)
                If headerValue IsNot Nothing AndAlso CStr(headerValue).Trim().Contains("Причина отсутствия") Then
                    Dim reasonValue As Object = GetCellValue(ws, r, col)
                    If reasonValue IsNot Nothing AndAlso Not String.IsNullOrEmpty(CStr(reasonValue).Trim()) Then
                        hasAbsenceReason = True
                    End If
                    Exit For
                End If
            Next

            ' Для выходных и дней с причиной отсутствия не корректируем
            If isWeekendDay OrElse hasAbsenceReason Then
                Continue For
            End If

            ' Получаем максимальное время обеда из графика работы
            Dim workSchedule As String = GetWorkScheduleForEmployee(ws, r)
            Dim maxLunchMinutes As Integer = ExtractLunchMinutesFromSchedule(workSchedule)

            ' Находим колонку "Прогулял" и читаем время
            Dim progulalHours As Double = 0
            For col As Integer = 1 To lastCol
                Dim headerValue As Object = GetCellValue(ws, 1, col)
                If headerValue IsNot Nothing Then
                    Dim headerText As String = CStr(headerValue).Trim()
                    If headerText.Contains("Находился вне здания") OrElse headerText.Contains("Прогулял") Then
                        Dim progulalCell As Excel.Range = CType(ws.Cells(r, col), Excel.Range)
                        progulalHours = ParseOvertimeHours(progulalCell)
                        Marshal.FinalReleaseComObject(progulalCell)
                        Exit For
                    End If
                End If
            Next

            ' Читаем текущее значение "Находился в здании"
            Dim timeCell As Excel.Range = CType(ws.Cells(r, COL_TIME), Excel.Range)
            Dim currentTimeHours As Double = ParseOvertimeHours(timeCell)

            ' 1. Если прогулял < обеда, вычитаем разницу
            Dim progulalMinutes As Integer = CInt(progulalHours * 60)
            If progulalMinutes < maxLunchMinutes Then
                Dim differenceMinutes As Integer = maxLunchMinutes - progulalMinutes
                Dim differenceHours As Double = differenceMinutes / 60.0
                currentTimeHours = currentTimeHours - differenceHours
            End If

            ' 2. Добавляем утреннюю и вечернюю переработку
            Dim startTimeVal As Object = GetCellValue(ws, r, COL_START_TIME)
            Dim endTimeVal As Object = GetCellValue(ws, r, COL_END_TIME)
            Dim startTimeStr As String = If(startTimeVal Is Nothing, "", CStr(startTimeVal).Trim())
            Dim endTimeStr As String = If(endTimeVal Is Nothing, "", CStr(endTimeVal).Trim())

            ' Проверяем, что есть фактическое время прихода и ухода
            If Not String.IsNullOrEmpty(startTimeStr) AndAlso Not startTimeStr.Contains("Нет входа") AndAlso
               Not String.IsNullOrEmpty(endTimeStr) AndAlso Not endTimeStr.Contains("Нет выход") Then

                ' Извлекаем время начала и конца из графика
                Dim scheduleStartTime As TimeSpan? = ExtractStartTimeFromSchedule(workSchedule)
                
                ' Проверяем, пятница ли это
                Dim isFriday As Boolean = False
                If TypeOf dateVal Is Date Then
                    isFriday = (CDate(dateVal).DayOfWeek = DayOfWeek.Friday)
                ElseIf TypeOf dateVal Is Double Then
                    isFriday = (Date.FromOADate(CDbl(dateVal)).DayOfWeek = DayOfWeek.Friday)
                End If
                
                Dim scheduleEndTime As TimeSpan? = ExtractEndTimeFromSchedule(workSchedule, isFriday)

                ' Парсим фактическое время прихода и ухода
                Dim actualStartTime As TimeSpan? = ParseTimeFromCellValue(startTimeVal)
                Dim actualEndTime As TimeSpan? = ParseTimeFromCellValue(endTimeVal)

                If scheduleStartTime.HasValue AndAlso actualStartTime.HasValue Then
                    ' Утренняя переработка: если пришел РАНЬШЕ → добавляем разницу
                    If actualStartTime.Value < scheduleStartTime.Value Then
                        Dim morningOvertime As TimeSpan = scheduleStartTime.Value - actualStartTime.Value
                        currentTimeHours += morningOvertime.TotalHours
                    End If
                End If

                If scheduleEndTime.HasValue AndAlso actualEndTime.HasValue Then
                    ' Вечерняя переработка: если ушел ПОЗЖЕ → добавляем разницу
                    If actualEndTime.Value > scheduleEndTime.Value Then
                        Dim eveningOvertime As TimeSpan = actualEndTime.Value - scheduleEndTime.Value
                        currentTimeHours += eveningOvertime.TotalHours
                    End If
                End If
            End If

            ' Если получилось отрицательное значение, устанавливаем 0
            If currentTimeHours < 0 Then
                currentTimeHours = 0
            End If

            ' Записываем новое значение в формате Ч:ММ
            Dim totalMinutes As Integer = CInt(currentTimeHours * 60)
            Dim resultHours As Integer = totalMinutes \ 60
            Dim resultMinutes As Integer = totalMinutes Mod 60

            Dim timeString As String = $"{resultHours}:{resultMinutes:D2}"

            timeCell.NumberFormat = "@"
            timeCell.Value = timeString
            timeCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

            Marshal.FinalReleaseComObject(timeCell)
        Next

        ' ЭТАП 3.5: Пересчет "Находился вне здания" с вычетом обеда
        lastRow = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        
        For r As Integer = 2 To lastRow
            ' Пропускаем строки ИТОГО
            Dim markerVal As Object = GetCellValue(ws, r, COL_MARKER)
            If StringEquals(markerVal, "ИТОГО") Then Continue For

            ' Проверяем, выходной ли день
            Dim dateVal As Object = GetCellValue(ws, r, COL_DATE)
            Dim isWeekendDay As Boolean = IsWeekend(dateVal)

            ' Проверяем, есть ли причина отсутствия
            Dim hasAbsenceReason As Boolean = False
            Dim lastCol As Integer = ws.Cells(r, ws.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
            For col As Integer = 1 To lastCol
                Dim headerValue As Object = GetCellValue(ws, 1, col)
                If headerValue IsNot Nothing AndAlso CStr(headerValue).Trim().Contains("Причина отсутствия") Then
                    Dim reasonValue As Object = GetCellValue(ws, r, col)
                    If reasonValue IsNot Nothing AndAlso Not String.IsNullOrEmpty(CStr(reasonValue).Trim()) Then
                        hasAbsenceReason = True
                    End If
                    Exit For
                End If
            Next

            ' Проверяем наличие входа и выхода
            Dim startTimeVal As Object = GetCellValue(ws, r, COL_START_TIME)
            Dim endTimeVal As Object = GetCellValue(ws, r, COL_END_TIME)
            Dim startTimeStr As String = If(startTimeVal Is Nothing, "", CStr(startTimeVal).Trim())
            Dim endTimeStr As String = If(endTimeVal Is Nothing, "", CStr(endTimeVal).Trim())
            
            Dim hasNoEntryOrExit As Boolean = String.IsNullOrEmpty(startTimeStr) OrElse 
                                              String.IsNullOrEmpty(endTimeStr) OrElse
                                              startTimeStr.Contains("Нет входа") OrElse 
                                              endTimeStr.Contains("Нет выход")

            ' Находим колонку "Прогулял" / "Находился вне здания"
            For col As Integer = 1 To lastCol
                Dim headerValue As Object = GetCellValue(ws, 1, col)
                If headerValue IsNot Nothing Then
                    Dim headerText As String = CStr(headerValue).Trim()
                    If headerText.Contains("Находился вне здания") OrElse headerText.Contains("Прогулял") Then
                        Dim progulalCell As Excel.Range = CType(ws.Cells(r, col), Excel.Range)
                        
                        ' Для выходных, дней с причиной отсутствия и дней без входа/выхода устанавливаем 0
                        If isWeekendDay OrElse hasAbsenceReason OrElse hasNoEntryOrExit Then
                            progulalCell.NumberFormat = "@"
                            progulalCell.Value = "0:00"
                            progulalCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                            Marshal.FinalReleaseComObject(progulalCell)
                            Exit For
                        End If
                        
                        ' Получаем максимальное время обеда из графика работы
                        Dim workSchedule As String = GetWorkScheduleForEmployee(ws, r)
                        Dim maxLunchMinutes As Integer = ExtractLunchMinutesFromSchedule(workSchedule)
                        
                        Dim progulalHours As Double = ParseOvertimeHours(progulalCell)
                        
                        ' Вычитаем обед
                        Dim progulalMinutes As Integer = CInt(progulalHours * 60)
                        Dim newProgulalMinutes As Integer = progulalMinutes - maxLunchMinutes
                        
                        ' Если меньше 0, то ставим 0
                        If newProgulalMinutes < 0 Then
                            newProgulalMinutes = 0
                        End If
                        
                        ' Записываем новое значение в формате Ч:ММ
                        Dim resultHours As Integer = newProgulalMinutes \ 60
                        Dim resultMinutes As Integer = newProgulalMinutes Mod 60
                        
                        Dim timeString As String = $"{resultHours}:{resultMinutes:D2}"
                        
                        progulalCell.NumberFormat = "@"
                        progulalCell.Value = timeString
                        progulalCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        
                        Marshal.FinalReleaseComObject(progulalCell)
                        Exit For
                    End If
                End If
            Next
        Next

        ' ЭТАП 4: Окрашивание и анализ времени
        lastRow = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row

        For r As Integer = lastRow To 2 Step -1
            Dim dateVal As Object = GetCellValue(ws, r, COL_DATE)
            Dim timeVal As Object = GetCellValue(ws, r, COL_TIME)

            ' ==================== ПРОВЕРКА НА ВЫХОДНОЙ ИЛИ ПРИЧИНУ ОТСУТСТВИЯ ====================
            ' Пропускаем строки ИТОГО
            If StringEquals(ws.Cells(r, COL_MARKER).Value2, "ИТОГО") Then
                Continue For
            End If

            Dim isWeekendDay As Boolean = IsWeekend(dateVal)
            Dim hasAbsenceReason As Boolean = False
            Dim lastCol As Integer = ws.Cells(r, ws.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

            ' Ищем колонку "Причина отсутствия" и проверяем, есть ли в ней значение
            For col As Integer = 1 To lastCol
                Dim headerValue As Object = GetCellValue(ws, 1, col)
                If headerValue IsNot Nothing Then
                    Dim headerText As String = CStr(headerValue).Trim()
                    If headerText.Contains("Причина отсутствия") Then
                        Dim reasonValue As Object = GetCellValue(ws, r, col)
                        If reasonValue IsNot Nothing AndAlso Not String.IsNullOrEmpty(CStr(reasonValue).Trim()) Then
                            hasAbsenceReason = True
                        End If
                        Exit For
                    End If
                End If
            Next

            'красим 0 проходы (не красим для выходных и не красим если есть причина отсутствия)
            If IsZeroTime(timeVal) AndAlso Not isWeekendDay AndAlso Not hasAbsenceReason Then
                Dim tcell As Excel.Range = CType(ws.Cells(r, COL_TIME), Excel.Range)
                tcell.Interior.Color = paleYellow ' фон желтый
                tcell.Font.Color = ColorTranslator.ToOle(Color.Red) ' шрифт красный
                Marshal.FinalReleaseComObject(tcell)
            End If

            ' ==================== ЕСЛИ ВЫХОДНОЙ ИЛИ ЕСТЬ ПРИЧИНА ОТСУТСТВИЯ ====================
            If isWeekendDay OrElse hasAbsenceReason Then
                ' Для выходных и дней с причиной отсутствия: переработка = конец дня - начало дня (БЕЗ обеда)
                Dim startTimeVal As Object = GetCellValue(ws, r, COL_START_TIME)
                Dim endTimeVal As Object = GetCellValue(ws, r, COL_END_TIME)
                Dim startTimeStr As String = If(startTimeVal Is Nothing, "", CStr(startTimeVal).Trim())
                Dim endTimeStr As String = If(endTimeVal Is Nothing, "", CStr(endTimeVal).Trim())

                If Not String.IsNullOrEmpty(startTimeStr) AndAlso Not startTimeStr.Contains("Нет входа") AndAlso
                   Not String.IsNullOrEmpty(endTimeStr) AndAlso Not endTimeStr.Contains("Нет выход") Then

                    Dim startTime As TimeSpan? = ParseTimeFromCellValue(startTimeVal)
                    Dim endTime As TimeSpan? = ParseTimeFromCellValue(endTimeVal)

                    If startTime.HasValue AndAlso endTime.HasValue Then
                        Dim workHours As Double = (endTime.Value - startTime.Value).TotalHours

                        ' Записываем в фактическую переработку в формате Ч:ММ
                        Dim totalMinutes As Integer = CInt(Math.Abs(workHours) * 60)
                        Dim resultHours As Integer = totalMinutes \ 60
                        Dim resultMinutes As Integer = totalMinutes Mod 60

                        Dim timeString As String
                        If workHours < 0 Then
                            timeString = $"-{resultHours}:{resultMinutes:D2}"
                        Else
                            timeString = $"{resultHours}:{resultMinutes:D2}"
                        End If

                        Dim overtimeCell As Excel.Range = CType(ws.Cells(r, COL_OVERTIME), Excel.Range)
                        overtimeCell.NumberFormat = "@"
                        overtimeCell.Value = timeString
                        overtimeCell.HorizontalAlignment = If(workHours >= 0, Excel.XlHAlign.xlHAlignRight, Excel.XlHAlign.xlHAlignLeft)
                        
                        ' Добавляем комментарий для выходного/причины отсутствия
                        Dim dateCell As Excel.Range = CType(ws.Cells(r, COL_DATE), Excel.Range)
                        Dim commentText As String = GetSimpleComment(dateCell.Value2, isWeekendDay, hasAbsenceReason)
                        AddSimpleComment(overtimeCell, commentText)
                        Marshal.FinalReleaseComObject(dateCell)
                        Marshal.FinalReleaseComObject(overtimeCell)
                    End If
                Else
                    ' Если нет отработанного времени, ставим 0:00
                    Dim overtimeCell As Excel.Range = CType(ws.Cells(r, COL_OVERTIME), Excel.Range)
                    overtimeCell.NumberFormat = "@"
                    overtimeCell.Value = "0:00"
                    overtimeCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    
                    ' Добавляем комментарий для выходного/причины отсутствия без отработанного времени
                    Dim dateCell As Excel.Range = CType(ws.Cells(r, COL_DATE), Excel.Range)
                    Dim commentText As String = GetSimpleComment(dateCell.Value2, isWeekendDay, hasAbsenceReason)
                    AddSimpleComment(overtimeCell, commentText)
                    Marshal.FinalReleaseComObject(dateCell)
                    Marshal.FinalReleaseComObject(overtimeCell)
                End If
            Else
                ' ==================== ОБЫЧНЫЙ РАБОЧИЙ ДЕНЬ БЕЗ ПРИЧИНЫ ОТСУТСТВИЯ ====================
                AnalyzeWorkTimeViolations(ws, r)

                ' Проверяем, есть ли вход и выход для сотрудника
                Dim startTimeVal As Object = GetCellValue(ws, r, COL_START_TIME)
                Dim endTimeVal As Object = GetCellValue(ws, r, COL_END_TIME)
                Dim startTimeStr As String = If(startTimeVal Is Nothing, "", CStr(startTimeVal).Trim())
                Dim endTimeStr As String = If(endTimeVal Is Nothing, "", CStr(endTimeVal).Trim())

                ' Если нет входа или выхода, ставим 0 в фактическую переработку
                If String.IsNullOrEmpty(startTimeStr) OrElse String.IsNullOrEmpty(endTimeStr) OrElse
                   startTimeStr.Contains("Нет входа") OrElse endTimeStr.Contains("Нет выход") Then
                    Dim overtimeCell As Excel.Range = CType(ws.Cells(r, COL_OVERTIME), Excel.Range)
                    overtimeCell.ClearFormats()
                    overtimeCell.NumberFormat = "@"
                    overtimeCell.Value = "0"
                    overtimeCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    Marshal.FinalReleaseComObject(overtimeCell)
                End If
            End If
        Next

        ' Нормализуем значения переработки перед суммированием ИТОГО
        NormalizeSummaryOvertimeValues(ws)

        ' ЭТАП 5: Суммирование для строк ИТОГО
        For r As Integer = lastRow To 2 Step -1

            ' --- Обработка "Фактическая переработка" ТОЛЬКО для строки ИТОГО (14-й столбец) ---
            Dim marker As Object = ws.Cells(r, COL_MARKER).Value2
            If Not IsNothing(marker) AndAlso String.Equals(CStr(marker), "ИТОГО", StringComparison.CurrentCultureIgnoreCase) Then

                ' Добавляем формулу для пересчета фактической переработки по сотруднику
                Dim ocell As Excel.Range = CType(ws.Cells(r, COL_OVERTIME), Excel.Range) ' 14-й столбец

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


                ' Проверяем, что есть строки для суммирования
                If formulaStartRow < r - 1 Then
                    ' Считаем сумму программно
                    Dim totalHours As Double = 0

                    For row As Integer = formulaStartRow To r - 1
                        Dim cellRange As Excel.Range = Nothing
                        Try
                            Dim startObj As Object = GetCellValue(ws, row, COL_START_TIME)
                            Dim endObj As Object = GetCellValue(ws, row, COL_END_TIME)
                            Dim startText As String = If(startObj Is Nothing, String.Empty, CStr(startObj))
                            Dim endText As String = If(endObj Is Nothing, String.Empty, CStr(endObj))

                            Dim forceZero As Boolean =
                                String.IsNullOrWhiteSpace(startText) OrElse
                                String.IsNullOrWhiteSpace(endText) OrElse
                                startText.IndexOf("нет вход", StringComparison.CurrentCultureIgnoreCase) >= 0 OrElse
                                endText.IndexOf("нет выход", StringComparison.CurrentCultureIgnoreCase) >= 0

                            Dim numericValue As Double
                            If forceZero Then
                                numericValue = 0
                            Else
                                cellRange = CType(ws.Cells(row, COL_OVERTIME), Excel.Range)
                                numericValue = ParseOvertimeHours(cellRange)
                            End If

                            totalHours += numericValue
                        Finally
                            If cellRange IsNot Nothing Then Marshal.FinalReleaseComObject(cellRange)
                        End Try
                    Next

                    ' ��⠭�������� ���᫥���� ���祭�� � ⥪�⮢�� �ଠ�
                    Dim absHours As Double = Math.Abs(totalHours)
                    Dim totalMinutes As Integer = CInt(absHours * 60.0R)
                    Dim resultHours As Integer = totalMinutes \ 60
                    Dim resultMinutes As Integer = totalMinutes Mod 60

                    If totalMinutes = 0 Then
                        ocell.Value2 = "0:00"
                    ElseIf totalHours < 0 Then
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
                        ocell.Interior.Color = ColorTranslator.ToOle(Color.LimeGreen)        ' > 0 → зелёный
                    ElseIf hours >= -1 Then
                        ocell.Interior.Color = ColorTranslator.ToOle(Color.Yellow)           ' -1..0 → жёлтый (напр. -0,27)
                    Else
                        ocell.Interior.Color = ColorTranslator.ToOle(Color.Red)              ' < -1 → красный (напр. -1,27 или -167)
                    End If
                End If

                Marshal.FinalReleaseComObject(ocell)

                ' --- Пересчет "Находился в здании" для строки ИТОГО (7-й столбец) ---
                Dim timeCell As Excel.Range = CType(ws.Cells(r, COL_TIME), Excel.Range)
                
                ' Считаем сумму "Находился в здании" программно
                Dim totalTimeHours As Double = 0
                
                For row As Integer = formulaStartRow To r - 1
                    Dim timeCellRow As Excel.Range = Nothing
                    Try
                        timeCellRow = CType(ws.Cells(row, COL_TIME), Excel.Range)
                        Dim timeHours As Double = ParseOvertimeHours(timeCellRow)
                        totalTimeHours += timeHours
                    Finally
                        If timeCellRow IsNot Nothing Then Marshal.FinalReleaseComObject(timeCellRow)
                    End Try
                Next
                
                ' Форматируем и записываем сумму
                Dim totalTimeMinutes As Integer = CInt(Math.Abs(totalTimeHours) * 60)
                Dim timeResultHours As Integer = totalTimeMinutes \ 60
                Dim timeResultMinutes As Integer = totalTimeMinutes Mod 60
                
                Dim timeString As String
                If totalTimeMinutes = 0 Then
                    timeString = "0:00"
                ElseIf totalTimeHours < 0 Then
                    timeString = $"-{timeResultHours}:{timeResultMinutes:D2}"
                Else
                    timeString = $"{timeResultHours}:{timeResultMinutes:D2}"
                End If
                
                timeCell.NumberFormat = "@"
                timeCell.Value = timeString
                timeCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                
                Marshal.FinalReleaseComObject(timeCell)

                ' --- Пересчет "Находился вне здания" для строки ИТОГО ---
                Dim progulalCol As Integer = 0
                Dim lastColSearch As Integer = ws.Cells(r, ws.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
                
                ' Находим колонку "Прогулял" / "Находился вне здания"
                For col As Integer = 1 To lastColSearch
                    Dim headerValue As Object = GetCellValue(ws, 1, col)
                    If headerValue IsNot Nothing Then
                        Dim headerText As String = CStr(headerValue).Trim()
                        If headerText.Contains("Находился вне здания") OrElse headerText.Contains("Прогулял") Then
                            progulalCol = col
                            Exit For
                        End If
                    End If
                Next
                
                ' Если нашли колонку, пересчитываем сумму
                If progulalCol > 0 Then
                    Dim progulalCell As Excel.Range = CType(ws.Cells(r, progulalCol), Excel.Range)
                    
                    ' Считаем сумму "Находился вне здания" программно
                    Dim totalProgulalHours As Double = 0
                    
                    For row As Integer = formulaStartRow To r - 1
                        Dim progulalCellRow As Excel.Range = Nothing
                        Try
                            progulalCellRow = CType(ws.Cells(row, progulalCol), Excel.Range)
                            Dim progulalHours As Double = ParseOvertimeHours(progulalCellRow)
                            totalProgulalHours += progulalHours
                        Finally
                            If progulalCellRow IsNot Nothing Then Marshal.FinalReleaseComObject(progulalCellRow)
                        End Try
                    Next
                    
                    ' Форматируем и записываем сумму
                    Dim totalProgulalMinutes As Integer = CInt(Math.Abs(totalProgulalHours) * 60)
                    Dim progulalResultHours As Integer = totalProgulalMinutes \ 60
                    Dim progulalResultMinutes As Integer = totalProgulalMinutes Mod 60
                    
                    Dim progulalTimeString As String
                    If totalProgulalMinutes = 0 Then
                        progulalTimeString = "0:00"
                    ElseIf totalProgulalHours < 0 Then
                        progulalTimeString = $"-{progulalResultHours}:{progulalResultMinutes:D2}"
                    Else
                        progulalTimeString = $"{progulalResultHours}:{progulalResultMinutes:D2}"
                    End If
                    
                    progulalCell.NumberFormat = "@"
                    progulalCell.Value = progulalTimeString
                    progulalCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    
                    Marshal.FinalReleaseComObject(progulalCell)
                End If
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
            ' Переименование заголовков теперь выполняется в функции RenameHeaders

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
            ' Удаляем "Работа в праздничные дни" (13-я колонка)
            Dim holidayCol As Excel.Range = CType(ws.Columns(13), Excel.Range)
            holidayCol.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
            Marshal.FinalReleaseComObject(holidayCol)

            ' Удаляем "Комм. причины отсутствия" (10-я колонка)
            Dim commCol As Excel.Range = CType(ws.Columns(10), Excel.Range)
            commCol.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
            Marshal.FinalReleaseComObject(commCol)

            ' Удаляем "Причины не выхода" (9-я колонка)
            Dim reasonCol As Excel.Range = CType(ws.Columns(9), Excel.Range)
            reasonCol.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
            Marshal.FinalReleaseComObject(reasonCol)

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

    Private Sub RemoveEmptyReportSheets(wb As Excel.Workbook)
        ' Удаляем пустые листы с именем "Отчет"
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

                If usedRange IsNot Nothing Then Marshal.FinalReleaseComObject(usedRange)

                If isEmpty Then
                    sheetsToDelete.Add(ws)
                End If
            End If
        Next

        ' Удаляем найденные пустые листы
        For Each ws As Excel.Worksheet In sheetsToDelete
            ws.Delete()
            Marshal.FinalReleaseComObject(ws)
        Next
    End Sub

    Private Sub SortSheetsAlphabetically(wb As Excel.Workbook)
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

    Private Function ParseOvertimeHours(cell As Excel.Range) As Double
        If cell Is Nothing Then Return 0

        Dim rawValue As Object = Nothing
        Dim textValue As String = String.Empty
        Dim numberFormat As String = String.Empty

        Try
            rawValue = cell.Value2
        Catch
            rawValue = Nothing
        End Try

        Try
            textValue = CStr(cell.Text)
        Catch
            textValue = String.Empty
        End Try

        Try
            numberFormat = CStr(cell.NumberFormat)
        Catch
            numberFormat = String.Empty
        End Try

        textValue = If(textValue, String.Empty).Trim()
        numberFormat = If(numberFormat, String.Empty)

        If textValue.Length = 0 OrElse textValue = "-" Then
            Return 0
        End If

        Dim hoursFromText As Double
        If TryParseTimeText(textValue, hoursFromText) Then
            Return hoursFromText
        End If

        If rawValue IsNot Nothing Then
            Dim rawString As String = Convert.ToString(rawValue, CultureInfo.CurrentCulture)

            If TypeOf rawValue Is Double OrElse TypeOf rawValue Is Single OrElse TypeOf rawValue Is Decimal Then
                Dim dbl As Double = CDbl(rawValue)
                If LooksLikeTimeFormat(numberFormat) Then
                    Return dbl * 24.0R
                End If
                Return dbl
            End If

            Dim numericFromRaw As Double
            If Double.TryParse(rawString, NumberStyles.Float, CultureInfo.CurrentCulture, numericFromRaw) Then
                Return numericFromRaw
            End If
            If Double.TryParse(rawString, NumberStyles.Float, CultureInfo.InvariantCulture, numericFromRaw) Then
                Return numericFromRaw
            End If
        End If

        Dim numericFromText As Double
        If Double.TryParse(textValue, NumberStyles.Float, CultureInfo.CurrentCulture, numericFromText) Then
            Return numericFromText
        End If
        If Double.TryParse(textValue, NumberStyles.Float, CultureInfo.InvariantCulture, numericFromText) Then
            Return numericFromText
        End If

        Return 0
    End Function

    Private Function TryParseTimeText(text As String, ByRef hours As Double) As Boolean
        hours = 0
        If String.IsNullOrWhiteSpace(text) Then Return False

        Dim s As String = text.Trim()
        Dim sign As Double = 1
        If s.StartsWith("-"c) Then
            sign = -1
            s = s.Substring(1)
        ElseIf s.StartsWith("+"c) Then
            s = s.Substring(1)
        End If
        s = s.Trim()

        Dim parts() As String = s.Split(":"c)
        If parts.Length < 2 Then
            Return False
        End If

        Dim hh As Integer
        Dim mm As Integer = 0
        Dim ss As Integer = 0

        If Not Integer.TryParse(parts(0), NumberStyles.Integer, CultureInfo.InvariantCulture, hh) Then
            Return False
        End If
        If parts.Length >= 2 AndAlso Not Integer.TryParse(parts(1), NumberStyles.Integer, CultureInfo.InvariantCulture, mm) Then
            Return False
        End If
        If parts.Length >= 3 Then Integer.TryParse(parts(2), NumberStyles.Integer, CultureInfo.InvariantCulture, ss)

        hours = sign * (hh + mm / 60.0R + ss / 3600.0R)
        Return True
    End Function

    Private Function LooksLikeTimeFormat(formatString As String) As Boolean
        If String.IsNullOrEmpty(formatString) Then Return False
        Dim nf As String = formatString.ToLower(CultureInfo.InvariantCulture)
        If nf.Contains("h") OrElse nf.Contains(":") Then Return True
        If nf.Contains("час") OrElse nf.Contains("мин") Then Return True
        Return False
    End Function
    Private Sub NormalizeSummaryOvertimeValues(ws As Excel.Worksheet)
        If ws Is Nothing Then Exit Sub

        Dim lastRow As Integer = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        For totalRow As Integer = ROW_DATA_START To lastRow
            Dim markerObj As Object = GetCellValue(ws, totalRow, COL_MARKER)
            If Not StringEquals(markerObj, "ИТОГО") Then Continue For

            Dim firstDataRow As Integer = totalRow - 1
            While firstDataRow > ROW_DATA_START AndAlso Not StringEquals(GetCellValue(ws, firstDataRow, COL_MARKER), "ИТОГО")
                firstDataRow -= 1
            End While

            If StringEquals(GetCellValue(ws, firstDataRow, COL_MARKER), "ИТОГО") Then
                firstDataRow += 1
            End If

            If firstDataRow < ROW_DATA_START Then
                firstDataRow = ROW_DATA_START
            End If

            Dim totalHours As Double = 0
            For dataRow As Integer = firstDataRow To totalRow - 1
                Dim cellRange As Excel.Range = Nothing
                Try
                    cellRange = CType(ws.Cells(dataRow, COL_OVERTIME), Excel.Range)
                    totalHours += ParseOvertimeHours(cellRange)
                Finally
                    If cellRange IsNot Nothing Then Marshal.FinalReleaseComObject(cellRange)
                End Try
            Next

            Dim targetCell As Excel.Range = CType(ws.Cells(totalRow, COL_OVERTIME), Excel.Range)
            Try
                Dim totalMinutes As Integer = CInt(Math.Round(Math.Abs(totalHours) * 60.0R, MidpointRounding.AwayFromZero))
                Dim resultHours As Integer = totalMinutes \ 60
                Dim resultMinutes As Integer = totalMinutes Mod 60

                If totalMinutes = 0 Then
                    targetCell.Value2 = "0:00"
                ElseIf totalHours < 0 Then
                    targetCell.Value2 = $"-{resultHours}:{resultMinutes:D2}"
                Else
                    targetCell.Value2 = $"{resultHours}:{resultMinutes:D2}"
                End If

                targetCell.NumberFormat = "@"

                Dim ov As Object = targetCell.Value2
                Dim hoursForColor As Double
                Dim haveHours As Boolean = False

                If IsNumeric(ov) Then
                    Dim v As Double = CDbl(ov)
                    Dim nf As String = CStr(targetCell.NumberFormat)
                    Dim isTimeFmt As Boolean = (nf.IndexOf(":", StringComparison.Ordinal) >= 0) OrElse
                                           (nf.IndexOf("[h", StringComparison.OrdinalIgnoreCase) >= 0)
                    hoursForColor = If(isTimeFmt, v * 24.0R, v)
                    haveHours = True
                Else
                    Dim s As String = CStr(ov).Trim()
                    Dim sign As Double = 1
                    If s.StartsWith("-"c) Then sign = -1 : s = s.Substring(1)
                    If s.StartsWith("+"c) Then s = s.Substring(1)
                    Dim parts() As String = s.Split(":"c)
                    If parts.Length >= 2 Then
                        Dim hh As Double, mm As Double, ss As Double
                        If Double.TryParse(parts(0), hh) AndAlso Double.TryParse(parts(1), mm) Then
                            If parts.Length >= 3 Then Double.TryParse(parts(2), ss)
                            hoursForColor = sign * (hh + mm / 60.0R + ss / 3600.0R)
                            haveHours = True
                        End If
                    End If
                End If

                If haveHours Then
                    targetCell.FormatConditions.Delete()

                    If hoursForColor > 0 Then
                        targetCell.Interior.Color = ColorTranslator.ToOle(Color.LimeGreen)
                    ElseIf hoursForColor >= -1 Then
                        targetCell.Interior.Color = ColorTranslator.ToOle(Color.Yellow)
                    Else
                        targetCell.Interior.Color = ColorTranslator.ToOle(Color.Red)
                    End If
                End If
            Finally
                Marshal.FinalReleaseComObject(targetCell)
            End Try
        Next
    End Sub

    Private Function GetCellValue(ws As Excel.Worksheet, row As Integer, col As Integer) As Object
        Dim rng As Excel.Range = CType(ws.Cells(row, col), Excel.Range)
        Dim v As Object = rng.Value2
        Marshal.FinalReleaseComObject(rng)
        Return v
    End Function

    Private Function NormalizeMarkerText(value As String) As String
        If String.IsNullOrEmpty(value) Then Return String.Empty

        Dim normalized As String = value

        normalized = normalized.Replace(ChrW(&HA0), " ")
        normalized = normalized.Replace(ChrW(&H202F), " ")
        normalized = normalized.Replace(ChrW(&H2007), " ")

        normalized = normalized.Trim()

        While normalized.Length > 0 AndAlso Char.IsPunctuation(normalized(normalized.Length - 1))
            normalized = normalized.Substring(0, normalized.Length - 1).TrimEnd()
        End While

        normalized = Regex.Replace(normalized, "\s+", " ")

        Return normalized
    End Function

    Private Function StringEquals(v As Object, expected As String) As Boolean
        If v Is Nothing Then Return False
        Dim actualNormalized As String = NormalizeMarkerText(CStr(v))
        Dim expectedNormalized As String = NormalizeMarkerText(expected)
        If expectedNormalized.Length = 0 Then Return actualNormalized.Length = 0
        Return actualNormalized.StartsWith(expectedNormalized, True, CultureInfo.CurrentCulture)
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

    ' Получает график работы для сотрудника из 16-й колонки
    Private Function GetWorkScheduleForEmployee(ws As Excel.Worksheet, row As Integer) As String
        ' Всегда читаем из 15-й колонки (колонка O)
        Const SCHEDULE_COLUMN As Integer = 16

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

    ' Извлекает время обеда из текста графика (формат "обед:60" или "обед 60")
    Private Function ExtractLunchMinutesFromSchedule(scheduleText As String) As Integer
        If String.IsNullOrEmpty(scheduleText) Then Return 45 ' По умолчанию 45 минут

        ' Ищем паттерн "обед:60", "обед 60", "обед: 60"
        Dim lunchPattern As String = "обед\s*:?\s*(\d+)"
        Dim match As Match = Regex.Match(scheduleText, lunchPattern, RegexOptions.IgnoreCase)

        If match.Success Then
            Dim minutes As Integer = Integer.Parse(match.Groups(1).Value)
            Return minutes
        End If

        Return 45 ' По умолчанию 45 минут
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

        ' Если не число, пробуем распарсить как текст: "8:30", "08:30", "8:30:00"
        ' Сначала пробуем с секундами
        Dim timePatternWithSeconds As String = "(\d{1,2})[:-](\d{2})[:-](\d{2})"
        Dim matchWithSeconds As Match = Regex.Match(timeValue, timePatternWithSeconds)

        If matchWithSeconds.Success Then
            Dim hour As Integer = Integer.Parse(matchWithSeconds.Groups(1).Value)
            Dim minute As Integer = Integer.Parse(matchWithSeconds.Groups(2).Value)
            Dim second As Integer = Integer.Parse(matchWithSeconds.Groups(3).Value)
            Return New TimeSpan(hour, minute, second)
        End If

        ' Если не получилось с секундами, пробуем без них
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

                    ' Настраиваем размер комментария
                    If startCell.Comment IsNot Nothing Then
                        startCell.Comment.Shape.Width = 300
                        startCell.Comment.Shape.Height = 100
                        startCell.Comment.Shape.TextFrame.AutoSize = True
                    End If
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

                    ' Настраиваем размер комментария
                    If endCell.Comment IsNot Nothing Then
                        endCell.Comment.Shape.Width = 300
                        endCell.Comment.Shape.Height = 100
                        endCell.Comment.Shape.TextFrame.AutoSize = True
                    End If
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

    Private Function ColumnIndexToLetter(columnIndex As Integer) As String
        Dim columnLetter As String = ""
        Dim temp As Integer

        While columnIndex > 0
            temp = (columnIndex - 1) Mod 26
            columnLetter = Chr(temp + 65) & columnLetter
            columnIndex = (columnIndex - temp - 1) \ 26
        End While

        Return columnLetter
    End Function

    Private Function GetSimpleComment(dateValue As Object, isWeekend As Boolean, hasAbsenceReason As Boolean) As String
        ' Если есть причина отсутствия - возвращаем соответствующий текст
        If hasAbsenceReason Then
            Return "По причине отсутствия"
        End If
        
        ' Если выходной - определяем день недели
        If isWeekend AndAlso dateValue IsNot Nothing Then
            Try
                Dim dateObj As Date
                If TypeOf dateValue Is Date Then
                    dateObj = CDate(dateValue)
                ElseIf TypeOf dateValue Is Double Then
                    dateObj = Date.FromOADate(CDbl(dateValue))
                Else
                    Date.TryParse(dateValue.ToString(), dateObj)
                End If
                
                Dim dayOfWeek As DayOfWeek = dateObj.DayOfWeek
                If dayOfWeek = DayOfWeek.Saturday Then
                    Return "Суббота"
                ElseIf dayOfWeek = DayOfWeek.Sunday Then
                    Return "Воскресенье"
                End If
            Catch
                ' Если не удалось определить дату, просто вернем "Выходной"
                Return "Выходной"
            End Try
        End If
        
        Return ""
    End Function

    Private Sub AddSimpleComment(cell As Excel.Range, commentText As String)
        If cell Is Nothing OrElse String.IsNullOrEmpty(commentText) Then Return
        
        Try
            ' Удаляем старый комментарий, если есть
            If cell.Comment IsNot Nothing Then
                cell.Comment.Delete()
            End If
            
            ' Добавляем новый комментарий
            Dim comment As Excel.Comment = cell.AddComment(commentText)
            If comment IsNot Nothing AndAlso comment.Shape IsNot Nothing Then
                comment.Shape.TextFrame.AutoSize = True
            End If
        Catch
            ' Игнорируем ошибки при работе с комментариями
        End Try
    End Sub

End Module

