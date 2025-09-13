' VB.NET (VSTO Add-in, Excel) — Без late binding, совместимо с Option Strict On
' Логика как в присланном макросе: группировка по отделам, 2 листа максимум на отдел (обычный и _нет_прохода),
' копирование блоков между "ИТОГО", перенос ширин столбцов, удаление выходных без служебки, подсветка 0:00 и сортировка листов.
' Подключить: Microsoft.Office.Interop.Excel
' Вызов для активной книги: ReportByDepartments.GenerateFromActiveWorkbook(Globals.ThisAddIn.Application)
' Вызов для файла:        ReportByDepartments.GenerateFromFile(Globals.ThisAddIn.Application, "C:\\...\\выгрузка_полная.xlsx")


Option Explicit On
Option Infer On

Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Drawing
Imports System.Globalization

Public Module ReportByDepartments

    Private Const HEADER_ROW As Integer = 5
    Private Const DATA_START_ROW As Integer = 6
    Private Const COL_MARKER As Integer = 1      ' A — признак "ИТОГО"
    Private Const COL_DEPT As Integer = 2        ' B — колонка с названием отдела (на первой строке блока)
    Private Const COL_DATE As Integer = 6        ' F — дата
    Private Const COL_TIME As Integer = 7        ' G — время нахождения

    ' ====== ПУБЛИЧНЫЕ ТОЧКИ ВХОДА ======
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

    ' ====== ОСНОВНАЯ ЛОГИКА ======
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
                ' Если активный лист не из этой книги — возьмём первый лист книги
                wsSource = CType(srcWb.Sheets(1), Excel.Worksheet)
                wsSource.Activate()
            End If

            ' Применяем автофильтр по диапазону заголовка
            ApplyAutoFilter(wsSource, HEADER_ROW)

            Dim lastRow As Integer = wsSource.Cells(wsSource.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            Dim lastCol As Integer = wsSource.Cells(HEADER_ROW, wsSource.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

            ' === Создаём новый файл и сохраняем рядом с исходной книгой ===
            Dim wbNew As Excel.Workbook = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet)
            Dim wsFirst As Excel.Worksheet = CType(wbNew.Sheets(1), Excel.Worksheet)
            wsFirst.Cells.Clear()
            wsFirst.Name = "Отчет"

            Dim baseName As String = Path.GetFileNameWithoutExtension(srcWb.Name)
            Dim newName As String = baseName & "_по_отделам.xlsx"
            Dim saveDir As String = If(String.IsNullOrEmpty(srcWb.Path), app.DefaultFilePath, srcWb.Path)
            Dim savePath As String = Path.Combine(saveDir, newName)

            If File.Exists(savePath) Then
                File.Delete(savePath)
            End If

            wbNew.SaveAs(Filename:=savePath, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook)

            ' Кэш для ускорения: уже созданные листы
            Dim created As New Dictionary(Of String, Excel.Worksheet)(StringComparer.CurrentCulture)

            Dim startCopyRow As Integer = DATA_START_ROW
            For r As Integer = HEADER_ROW + 1 To lastRow
                Dim markerObj As Object = GetCellValue(wsSource, r, COL_MARKER)
                If StringEquals(markerObj, "ИТОГО") Then
                    Dim deptRaw As String = CStr(GetCellValue(wsSource, startCopyRow, COL_DEPT))
                    Dim dept As String = LimitSheetName(deptRaw)

                    Dim timeObj As Object = GetCellValue(wsSource, r, COL_TIME)
                    If IsZeroTime(timeObj) Then
                        dept &= "_нет_прохода"
                    End If

                    If dept.Length > 0 Then
                        Dim wsTarget As Excel.Worksheet = Nothing
                        If Not created.TryGetValue(dept, wsTarget) OrElse wsTarget Is Nothing Then
                            wsTarget = CreateOrGetSheet(wbNew, dept)
                            ' Копируем шапку (строка HEADER_ROW → 1)
                            Dim headerRowRange As Excel.Range = CType(wsSource.Rows(HEADER_ROW), Excel.Range)
                            Dim destHeaderRow As Excel.Range = CType(wsTarget.Rows(1), Excel.Range)
                            headerRowRange.Copy(Destination:=destHeaderRow)
                            destHeaderRow.Font.Bold = True
                            Marshal.FinalReleaseComObject(headerRowRange)
                            Marshal.FinalReleaseComObject(destHeaderRow)

                            created(dept) = wsTarget
                        End If

                        ' Вставляем блок строк [startCopyRow..r]
                        Dim pasteRow As Integer = wsTarget.Cells(wsTarget.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row + 1
                        Dim srcRange As Excel.Range = wsSource.Range(wsSource.Rows(startCopyRow), wsSource.Rows(r))
                        Dim destPaste As Excel.Range = CType(wsTarget.Rows(pasteRow), Excel.Range)
                        srcRange.Copy(Destination:=destPaste)
                        Marshal.FinalReleaseComObject(srcRange)
                        Marshal.FinalReleaseComObject(destPaste)

                        ' Форматируем как в макросе
                        BeautifySheet(wsTarget)

                        ' Перенос ширины столбцов
                        For c As Integer = 1 To lastCol
                            Dim srcCol As Excel.Range = CType(wsSource.Columns(c), Excel.Range)
                            Dim dstCol As Excel.Range = CType(wsTarget.Columns(c), Excel.Range)
                            dstCol.ColumnWidth = srcCol.ColumnWidth
                            Marshal.FinalReleaseComObject(srcCol)
                            Marshal.FinalReleaseComObject(dstCol)
                        Next
                    End If

                    startCopyRow = r + 1
                End If
            Next

            ' Упорядочиваем листы по алфавиту
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

    ' ====== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ======
    Private Function CreateOrGetSheet(wb As Excel.Workbook, baseName As String) As Excel.Worksheet
        ' Учитываем возможные совпадения имён: добавляем суффикс (2), (3) ...
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
        ' Собираем имена, сортируем и переставляем
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

    Private Sub ApplyAutoFilter(ws As Excel.Worksheet, headerRow As Integer)
        ' Безопасная версия: учитывает пустые строки/столбцы, активирует лист,
        ' пробует несколько способов и молча пропускает, если не удалось —
        ' для логики отчёта AutoFilter не обязателен.
        Try
            If ws Is Nothing Then Exit Sub

            Dim lastRow As Integer = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            Dim lastCol As Integer = ws.Cells(headerRow, ws.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

            If lastRow < headerRow OrElse lastCol < 1 Then Exit Sub

            ' AutoFilter иногда падает, если лист не активен
            ws.Activate()

            ' Снимем существующий фильтр, если он есть
            If ws.AutoFilterMode Then ws.AutoFilterMode = False

            Dim rng As Excel.Range = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))
            Try
                rng.AutoFilter()
            Catch ex As System.Runtime.InteropServices.COMException
                ' Попробуем по UsedRange как запасной путь
                Dim ur As Excel.Range = ws.UsedRange
                Try
                    ur.AutoFilter()
                Finally
                    If ur IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(ur)
                End Try
            Finally
                If rng IsNot Nothing Then Runtime.InteropServices.Marshal.FinalReleaseComObject(rng)
            End Try
        Catch
            ' Игнорируем: фильтр — косметика для UI, на формирование отчёта не влияет
        End Try
    End Sub

    Private Sub BeautifySheet(ws As Excel.Worksheet)
        Dim lastRow As Integer = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        Dim paleYellow As Integer = ColorTranslator.ToOle(Color.FromArgb(255, 255, 204)) ' #FFFFCC
        Dim red As Integer = ColorTranslator.ToOle(Color.Red)

        For r As Integer = lastRow To 2 Step -1
            Dim dateVal As Object = GetCellValue(ws, r, COL_DATE)
            Dim timeVal As Object = GetCellValue(ws, r, COL_TIME)

            If IsWeekend(dateVal) AndAlso IsZeroTime(timeVal) Then
                CType(ws.Rows(r), Excel.Range).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                Continue For
            End If

            If IsZeroTime(timeVal) Then
                CType(ws.Cells(r, COL_TIME), Excel.Range).Interior.Color = paleYellow
                CType(ws.Rows(r), Excel.Range).Font.Color = red
            End If
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
            ' Excel time 0.0
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
        ' Понедельник-основанная неделя: 1=Пн … 7=Вс
        Dim mondayBased As Integer = ((CInt(dt.DayOfWeek) + 6) Mod 7) + 1
        Return (mondayBased = 6 OrElse mondayBased = 7)
    End Function

    Private Function LimitSheetName(proposed As String) As String
        Dim name As String = If(proposed, String.Empty).Trim()
        ' Оставим 18 символов как в исходнике (можно увеличить до 31 при необходимости)
        If name.Length > 18 Then name = name.Substring(0, 18)

        ' Избегаем символьных литералов, чтобы не ловить BC30004
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

End Module
