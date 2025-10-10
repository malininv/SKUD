Option Strict Off
Option Explicit On
Option Infer On

Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Public Module WorkScheduleFiller
    Private Const HEADER_ROW As Integer = 5
    Private Const DATA_START_ROW As Integer = 6
    Private Const COL_EMPLOYEE As Integer = 3
    Private Const COL_DATE As Integer = 6

    Private ReadOnly DateFormats As String() = {"d.M.yyyy", "d.MM.yyyy", "dd.M.yyyy", "dd.MM.yyyy"}

    Public Function ApplyWorkSchedules(app As Excel.Application, schedulesPath As String) As Integer
        If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))
        If String.IsNullOrWhiteSpace(schedulesPath) OrElse Not File.Exists(schedulesPath) Then
            Throw New FileNotFoundException("Файл с графиками работы не найден.", schedulesPath)
        End If

        Dim targetWb As Excel.Workbook = app.ActiveWorkbook
        If targetWb Is Nothing Then Throw New InvalidOperationException("Нет активной книги с выгрузкой.")

        Dim targetWs As Excel.Worksheet = TryCast(app.ActiveSheet, Excel.Worksheet)
        If targetWs Is Nothing OrElse Not ReferenceEquals(targetWs.Parent, targetWb) Then
            targetWs = CType(targetWb.Sheets(1), Excel.Worksheet)
            targetWs.Activate()
        End If

        Dim calcPrev = app.Calculation
        Dim screenPrev = app.ScreenUpdating
        Dim eventsPrev = app.EnableEvents
        Dim alertsPrev = app.DisplayAlerts

        Dim schedulesWb As Excel.Workbook = Nothing
        Dim schedulesWs As Excel.Worksheet = Nothing
        Dim openedSchedules As Boolean = False

        Try
            app.Calculation = Excel.XlCalculation.xlCalculationManual
            app.ScreenUpdating = False
            app.EnableEvents = False
            app.DisplayAlerts = False

            schedulesWb = app.Workbooks.Open(schedulesPath, ReadOnly:=True)
            openedSchedules = True
            schedulesWs = CType(schedulesWb.Sheets(1), Excel.Worksheet)

            Dim workSchedules = LoadWorkSchedules(schedulesWs)

            Dim lastRow As Integer = targetWs.Cells(targetWs.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            If lastRow < DATA_START_ROW Then Return 0

            Dim reasonColumn As Integer = EnsureReasonColumn(targetWs)

            Dim reasonRange As Excel.Range = targetWs.Range(targetWs.Cells(DATA_START_ROW, reasonColumn), targetWs.Cells(lastRow, reasonColumn))
            reasonRange.ClearContents()
            Marshal.FinalReleaseComObject(reasonRange)

            Dim matches As Integer = 0

            For r As Integer = DATA_START_ROW To lastRow
                Dim fio As String = ReadString(targetWs, r, COL_EMPLOYEE)
                If String.IsNullOrEmpty(fio) Then Continue For

                Dim dayValue As Date? = ReadDate(targetWs, r, COL_DATE)
                If Not dayValue.HasValue Then Continue For

                Dim reason As String = FindReasonForDate(fio, dayValue.Value, workSchedules)
                If Not String.IsNullOrEmpty(reason) Then
                    WriteString(targetWs, r, reasonColumn, reason)
                    matches += 1
                End If
            Next

            Return matches
        Finally
            If openedSchedules AndAlso schedulesWb IsNot Nothing Then schedulesWb.Close(SaveChanges:=False)
            If schedulesWs IsNot Nothing Then Marshal.FinalReleaseComObject(schedulesWs)
            If schedulesWb IsNot Nothing Then Marshal.FinalReleaseComObject(schedulesWb)

            app.Calculation = calcPrev
            app.ScreenUpdating = screenPrev
            app.EnableEvents = eventsPrev
            app.DisplayAlerts = alertsPrev
        End Try
    End Function

    Private Function LoadWorkSchedules(ws As Excel.Worksheet) As Dictionary(Of String, WorkScheduleInfo)
        ' ==================== КОНСТАНТЫ ДЛЯ ФАЙЛА ГРАФИКОВ РАБОТЫ ====================
        ' Структура файла "графики работы.XLSX":
        ' - ФИО сотрудников находятся в колонке C (3), начиная со строки 7
        ' - Графики работы находятся в колонке F (6), начиная со строки 7
        ' - Пример: строка 7: C7="Алексеева Надежда Ивановна", F7="!! основной рабочий график с 8:00-17:00 (15:45 обед 12:15)"

        ' Используем константы из ReportByDepartments

        ' ==================== ОСНОВНАЯ ЛОГИКА ====================
        Dim result As New Dictionary(Of String, WorkScheduleInfo)(StringComparer.CurrentCultureIgnoreCase)
        Dim lastRow As Integer = ws.Cells(ws.Rows.Count, ReportByDepartments.SCHEDULES_FILE_SEARCH_COL).End(Excel.XlDirection.xlUp).Row
        If lastRow < ReportByDepartments.SCHEDULES_FILE_DATA_START_ROW Then Return result

        For r As Integer = ReportByDepartments.SCHEDULES_FILE_DATA_START_ROW To lastRow
            Dim fio As String = ReadString(ws, r, ReportByDepartments.SCHEDULES_FILE_COL_EMPLOYEE)
            If String.IsNullOrEmpty(fio) Then Continue For

            Dim scheduleText As String = ReadString(ws, r, ReportByDepartments.SCHEDULES_FILE_COL_SCHEDULE)
            If String.IsNullOrEmpty(scheduleText) Then Continue For

            Dim scheduleInfo As WorkScheduleInfo = ParseWorkSchedule(scheduleText)
            If scheduleInfo IsNot Nothing Then
                result(fio) = scheduleInfo
            End If
        Next

        Return result
    End Function

    Private Function ParseWorkSchedule(scheduleText As String) As WorkScheduleInfo
        If String.IsNullOrEmpty(scheduleText) Then Return Nothing

        ' Просто возвращаем объект с оригинальным текстом, без парсинга времени
        Return New WorkScheduleInfo With {
            .WorkStart = TimeSpan.Zero,      ' Не используется
            .WorkEnd = TimeSpan.Zero,        ' Не используется
            .LunchStart = Nothing,           ' Не используется
            .LunchEnd = Nothing,             ' Не используется
            .OriginalText = scheduleText.Trim() ' Сохраняем оригинальный текст
        }
    End Function

    Private Function FindReasonForDate(fio As String, dayValue As Date, schedules As Dictionary(Of String, WorkScheduleInfo)) As String
        Dim schedule As WorkScheduleInfo = Nothing
        If Not schedules.TryGetValue(fio, schedule) Then Return String.Empty

        ' Возвращаем оригинальный текст графика работы для всех дней
        Return schedule.OriginalText
    End Function


    Private Function EnsureReasonColumn(ws As Excel.Worksheet) As Integer
        Dim lastHeaderCol As Integer = ws.Cells(HEADER_ROW, ws.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
        For col As Integer = 1 To lastHeaderCol
            Dim caption As String = ReadString(ws, HEADER_ROW, col)
            If String.Compare(caption, ReportByDepartments.WORK_SCHEDULE_HEADER, True, CultureInfo.CurrentCulture) = 0 Then
                Return col
            End If
        Next

        Dim newCol As Integer = lastHeaderCol + 1
        WriteString(ws, HEADER_ROW, newCol, ReportByDepartments.WORK_SCHEDULE_HEADER)
        Dim headerCell As Excel.Range = CType(ws.Cells(HEADER_ROW, newCol), Excel.Range)
        headerCell.EntireColumn.NumberFormat = "@"
        Marshal.FinalReleaseComObject(headerCell)
        Return newCol
    End Function

    Private Function ReadString(ws As Excel.Worksheet, row As Integer, col As Integer) As String
        Dim cell As Excel.Range = CType(ws.Cells(row, col), Excel.Range)
        Try
            Dim value As Object = cell.Value2
            If value Is Nothing Then Return String.Empty
            Return value.ToString().Trim()
        Finally
            Marshal.FinalReleaseComObject(cell)
        End Try
    End Function

    Private Function ReadDate(ws As Excel.Worksheet, row As Integer, col As Integer) As Date?
        Dim cell As Excel.Range = CType(ws.Cells(row, col), Excel.Range)
        Try
            Dim value As Object = cell.Value2
            Return ParseDate(value)
        Finally
            Marshal.FinalReleaseComObject(cell)
        End Try
    End Function

    Private Sub WriteString(ws As Excel.Worksheet, row As Integer, col As Integer, value As String)
        Dim cell As Excel.Range = CType(ws.Cells(row, col), Excel.Range)
        Try
            cell.Value2 = value
        Finally
            Marshal.FinalReleaseComObject(cell)
        End Try
    End Sub

    Private Function ParseDate(raw As Object) As Date?
        If raw Is Nothing Then Return Nothing
        If TypeOf raw Is Double Then Return Date.FromOADate(CDbl(raw))

        Dim text As String = raw.ToString().Trim()
        If text.Length = 0 Then Return Nothing

        Dim dt As Date
        If Date.TryParseExact(text, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then Return dt
        If Date.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, dt) Then Return dt
        Return Nothing
    End Function

    Private Class WorkScheduleInfo
        Public Property WorkStart As TimeSpan
        Public Property WorkEnd As TimeSpan
        Public Property LunchStart As TimeSpan?
        Public Property LunchEnd As TimeSpan?
        Public Property OriginalText As String
    End Class
End Module
