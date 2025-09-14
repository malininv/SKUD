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

Public Module ReportByDepartments

    Private Const HEADER_ROW As Integer = 5
    Private Const DATA_START_ROW As Integer = 6
    Private Const COL_MARKER As Integer = 1 ' A — "ИТОГО"
    Private Const COL_DEPT As Integer = 2   ' B — отдел (первая строка блока)
    Private Const COL_DATE As Integer = 6   ' F — дата
    Private Const COL_TIME As Integer = 7   ' G — время

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

            ApplyAutoFilter(wsSource, HEADER_ROW)

            Dim lastRow As Integer = wsSource.Cells(wsSource.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            Dim lastCol As Integer = wsSource.Cells(HEADER_ROW, wsSource.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

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
            Dim startCopyRow As Integer = DATA_START_ROW

            For r As Integer = HEADER_ROW + 1 To lastRow
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
                            Dim headerRowRange As Excel.Range = CType(wsSource.Rows(HEADER_ROW), Excel.Range)
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

            ApplyAutoFilter(wsSource, HEADER_ROW)

            Dim lastRow As Integer = wsSource.Cells(wsSource.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            Dim lastCol As Integer = wsSource.Cells(HEADER_ROW, wsSource.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

            Dim saveRoot As String = If(String.IsNullOrEmpty(srcWb.Path), app.DefaultFilePath, srcWb.Path)
            Dim baseNamePerDept As String = Path.GetFileNameWithoutExtension(srcWb.Name)
            Dim targetDir As String = Path.Combine(saveRoot, LimitFileBaseName(baseNamePerDept) & "_по_отделам")
            If Not Directory.Exists(targetDir) Then Directory.CreateDirectory(targetDir)

            Dim deptToWb As New Dictionary(Of String, Excel.Workbook)(StringComparer.CurrentCulture)
            Dim deptToPath As New Dictionary(Of String, String)(StringComparer.CurrentCulture)
            Dim createdSheets As New Dictionary(Of String, Excel.Worksheet)(StringComparer.CurrentCulture)
            Dim deptHasAnySheet As New HashSet(Of String)(StringComparer.CurrentCulture)

            Dim startCopyRow As Integer = DATA_START_ROW
            For r As Integer = HEADER_ROW + 1 To lastRow
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
                        Dim headerRowRange As Excel.Range = CType(wsSource.Rows(HEADER_ROW), Excel.Range)
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
                    For c As Integer = 1 To lastCol
                        Dim srcCol As Excel.Range = CType(wsSource.Columns(c), Excel.Range)
                        Dim dstCol As Excel.Range = CType(wsTarget.Columns(c), Excel.Range)
                        dstCol.ColumnWidth = srcCol.ColumnWidth
                        Marshal.FinalReleaseComObject(srcCol)
                        Marshal.FinalReleaseComObject(dstCol)
                    Next

                    ' Save dept workbook incrementally
                    wbDept.Save()

                    startCopyRow = r + 1
                End If
            Next

            ' Finalize: sort sheets and close dept workbooks
            For Each kv In deptToWb
                Dim wbDept As Excel.Workbook = kv.Value
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
        For r As Integer = lastRow To 2 Step -1
            Dim dateVal As Object = GetCellValue(ws, r, COL_DATE)
            Dim timeVal As Object = GetCellValue(ws, r, COL_TIME)
            If IsWeekend(dateVal) AndAlso IsZeroTime(timeVal) Then
                CType(ws.Rows(r), Excel.Range).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                Continue For
            End If
            If IsZeroTime(timeVal) Then
                Dim tcell As Excel.Range = CType(ws.Cells(r, COL_TIME), Excel.Range)
                tcell.Interior.Color = red ' фон красный
                Marshal.FinalReleaseComObject(tcell)
            End If
        Next
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
        If name.Length > 31 Then name = name.Substring(0, 31)
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

End Module
