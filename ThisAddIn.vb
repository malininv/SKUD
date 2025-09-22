Option Strict On
Option Explicit On
Option Infer On

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ' Надстройка загрузилась — ничего дополнительно не требуется.
        ' Здесь можно включить логи/диагностику при желании.
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ' Очистка ресурсов при выгрузке надстройки (если понадобится).
    End Sub

    ' ====================== Публичные методы под 4 сценария ======================

    ' 1) Отделы по листам — из активной книги
    ' Возвращает полный путь к созданному файлу "<имя>_по_отделам.xlsx"
    Public Function RunSheetsActive() As String
        Return ReportByDepartments.GenerateFromActiveWorkbook(Me.Application)
    End Function

    ' 2) Отделы по листам — выбрать файл
    Public Function RunSheetsFromFile(filePath As String) As String
        Return ReportByDepartments.GenerateFromFile(Me.Application, filePath)
    End Function

    ' 3) Отделы по файлам — из активной книги
    ' Возвращает список путей к созданным файлам в подпапке "<имя>_по_отделам\"
    Public Function RunFilesPerDeptActive() As List(Of String)
        Return ReportByDepartments.GenerateFromActiveWorkbookPerDept(Me.Application)
    End Function

    ' 4) Отделы по файлам — выбрать файл
    Public Function RunFilesPerDeptFromFile(filePath As String) As List(Of String)
        Return ReportByDepartments.GenerateFromFilePerDept(Me.Application, filePath)
    End Function

    Public Function ApplyVacationsFromFile(leavesPath As String) As Integer
        Return LeaveReasonFiller.ApplyLeaveReasons(Me.Application, leavesPath)
    End Function

End Class