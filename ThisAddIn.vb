Option Strict On
Option Explicit On
Option Infer On

Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports System.IO

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ' Здесь ничего не обязательно. Надстройка просто загрузится.
        ' Кнопки на ленте будут вызывать публичные методы ниже.
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ' Очистка ресурсов если потребуется.
    End Sub

    ' === ПУБЛИЧНЫЕ МЕТОДЫ ДЛЯ ЛЕНТЫ / ВЫЗОВА ИЗ КОДА ===

    ' 1) Сформировать отчёт из АКТИВНОЙ книги (активный лист или первый лист)
    Public Sub RunReportForActiveWorkbook()
        Try
            Dim savedPath As String = ReportByDepartments.GenerateFromActiveWorkbook(Me.Application)
            Dim fname As String = If(String.IsNullOrEmpty(savedPath), "<неизвестно>", Path.GetFileName(savedPath))
            MessageBox.Show($"Готово: файл '{fname}' сохранён рядом с исходной книгой.",
                        "Отчёт по отделам", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Ошибка отчёта", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub RunReportForFile()
        Using dlg As New OpenFileDialog()
            dlg.Title = "Выберите файл Excel"
            dlg.Filter = "Excel книги (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
            dlg.Multiselect = False
            If dlg.ShowDialog() = DialogResult.OK Then
                Try
                    Dim savedPath As String = ReportByDepartments.GenerateFromFile(Me.Application, dlg.FileName)
                    Dim fname As String = If(String.IsNullOrEmpty(savedPath), "<неизвестно>", Path.GetFileName(savedPath))
                    MessageBox.Show($"Готово: файл '{fname}' сохранён рядом с исходной книгой.",
                                "Отчёт по отделам", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Ошибка отчёта", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

End Class