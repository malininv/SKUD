Option Strict Off
Option Explicit On
Option Infer On

Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Globalization
Imports System.Runtime.InteropServices

Public Module LeaveReasonFiller
    Private Const HEADER_ROW As Integer = 5
    Private Const DATA_START_ROW As Integer = 6
    Private Const COL_EMPLOYEE As Integer = 3
    Private Const COL_DATE As Integer = 6

    Private ReadOnly DateFormats As String() = {"d.M.yyyy", "d.MM.yyyy", "dd.M.yyyy", "dd.MM.yyyy"}

    Public Function ApplyLeaveReasons(app As Excel.Application, leavesPath As String) As Integer
        If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))
        If String.IsNullOrWhiteSpace(leavesPath) OrElse Not File.Exists(leavesPath) Then
            Throw New FileNotFoundException("Файл с отпусками не найден.", leavesPath)
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

        Dim leavesWb As Excel.Workbook = Nothing
        Dim leavesWs As Excel.Worksheet = Nothing
        Dim openedLeaves As Boolean = False

        Try
            app.Calculation = Excel.XlCalculation.xlCalculationManual
            app.ScreenUpdating = False
            app.EnableEvents = False
            app.DisplayAlerts = False

            leavesWb = app.Workbooks.Open(leavesPath, ReadOnly:=True)
            openedLeaves = True
            leavesWs = CType(leavesWb.Sheets(1), Excel.Worksheet)

            Dim intervals = LoadIntervals(leavesWs)

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

                Dim reason As String = FindReasonForDate(fio, dayValue.Value, intervals)
                If Not String.IsNullOrEmpty(reason) Then
                    WriteString(targetWs, r, reasonColumn, reason)
                    matches += 1
                End If
            Next

            Return matches
        Finally
            If openedLeaves AndAlso leavesWb IsNot Nothing Then leavesWb.Close(SaveChanges:=False)
            If leavesWs IsNot Nothing Then Marshal.FinalReleaseComObject(leavesWs)
            If leavesWb IsNot Nothing Then Marshal.FinalReleaseComObject(leavesWb)

            app.Calculation = calcPrev
            app.ScreenUpdating = screenPrev
            app.EnableEvents = eventsPrev
            app.DisplayAlerts = alertsPrev
        End Try
    End Function

    Private Function LoadIntervals(ws As Excel.Worksheet) As Dictionary(Of String, List(Of LeaveInterval))
        Dim result As New Dictionary(Of String, List(Of LeaveInterval))(StringComparer.CurrentCultureIgnoreCase)
        Dim lastRow As Integer = ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        If lastRow < 2 Then Return result

        For r As Integer = 2 To lastRow
            Dim fio As String = ReadString(ws, r, 1)
            If String.IsNullOrEmpty(fio) Then Continue For

            Dim startDate? As Date = ReadDate(ws, r, 5)
            Dim endDate? As Date = ReadDate(ws, r, 6)
            Dim reason As String = ReadString(ws, r, 7)

            If Not startDate.HasValue OrElse Not endDate.HasValue OrElse String.IsNullOrEmpty(reason) Then Continue For

            Dim normalizedStart As Date = startDate.Value.Date
            Dim normalizedEnd As Date = endDate.Value.Date
            If normalizedEnd < normalizedStart Then
                Dim tmp As Date = normalizedStart
                normalizedStart = normalizedEnd
                normalizedEnd = tmp
            End If

            Dim list As List(Of LeaveInterval) = Nothing
            If Not result.TryGetValue(fio, list) Then
                list = New List(Of LeaveInterval)()
                result(fio) = list
            End If

            list.Add(New LeaveInterval With {.StartDate = normalizedStart, .EndDate = normalizedEnd, .Reason = reason})
        Next

        For Each entry In result
            entry.Value.Sort(Function(a, b) a.StartDate.CompareTo(b.StartDate))
        Next

        Return result
    End Function

    Private Function FindReasonForDate(fio As String, dayValue As Date, data As Dictionary(Of String, List(Of LeaveInterval))) As String
        Dim intervals As List(Of LeaveInterval) = Nothing
        If Not data.TryGetValue(fio, intervals) Then Return String.Empty

        Dim current As Date = dayValue.Date
        For Each interval In intervals
            If current >= interval.StartDate AndAlso current <= interval.EndDate Then
                Return interval.Reason
            End If
        Next

        Return String.Empty
    End Function

    Private Function EnsureReasonColumn(ws As Excel.Worksheet) As Integer
        Dim lastHeaderCol As Integer = ws.Cells(HEADER_ROW, ws.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
        For col As Integer = 1 To lastHeaderCol
            Dim caption As String = ReadString(ws, HEADER_ROW, col)
            If String.Compare(caption, "Причина отсутствия", True, CultureInfo.CurrentCulture) = 0 Then
                Return col
            End If
        Next

        Dim newCol As Integer = lastHeaderCol + 1
        WriteString(ws, HEADER_ROW, newCol, "Причина отсутствия")
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

    Public Function ReadDate(ws As Excel.Worksheet, row As Integer, col As Integer) As Date?
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

    Private Class LeaveInterval
        Public Property StartDate As Date
        Public Property EndDate As Date
        Public Property Reason As String
    End Class
End Module
