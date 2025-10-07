Option Strict On
Option Explicit On
Option Infer On

Imports Microsoft.Office.Tools.Ribbon
Imports System.Windows.Forms
Imports System.IO
Imports System.Diagnostics

Public Class Ribbon1

    Private Sub Ribbon1_Load(sender As Object, e As RibbonUIEventArgs) Handles MyBase.Load
        ' Здесь можно скрывать/показывать элементы при старте, не обязательно.
    End Sub

    ' === 1) ОТДЕЛЫ ПО ЛИСТАМ — из активной книги ===
    Private Sub btnSheetsActive_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSheetsActive.Click
        Try
            Dim savedPath As String = Globals.ThisAddIn.RunSheetsActive()
            If Not String.IsNullOrEmpty(savedPath) Then
                Dim fname As String = Path.GetFileName(savedPath)
                MessageBox.Show($"Готово: файл '{fname}' сохранён рядом с исходной книгой.",
                                "Отделы по листам (активная)", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Try
                    Process.Start("explorer.exe", "/select,""" & savedPath & """")
                Catch
                End Try
            Else
                MessageBox.Show("Файл не был сохранён.", "Отделы по листам (активная)",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    ' === 2) ОТДЕЛЫ ПО ЛИСТАМ — выбрать файл ===
    Private Sub btnSheetsFile_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSheetsFile.Click
        Using dlg As New OpenFileDialog()
            dlg.Title = "Выберите файл Excel"
            dlg.Filter = "Excel книги (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
            dlg.Multiselect = False
            If dlg.ShowDialog() = DialogResult.OK Then
                Try
                    Dim savedPath As String = Globals.ThisAddIn.RunSheetsFromFile(dlg.FileName)
                    If Not String.IsNullOrEmpty(savedPath) Then
                        Dim fname As String = Path.GetFileName(savedPath)
                        MessageBox.Show($"Готово: файл '{fname}' сохранён рядом с исходной книгой.",
                                        "Отделы по листам (файл)", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Try
                            Process.Start("explorer.exe", "/select,""" & savedPath & """")
                        Catch
                        End Try
                    Else
                        MessageBox.Show("Файл не был сохранён.", "Отделы по листам (файл)",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                End Try
            End If
        End Using
    End Sub

    ' === 3) ОТДЕЛЫ ПО ФАЙЛАМ — из активной книги ===
    Private Sub btnFilesActive_Click(sender As Object, e As RibbonControlEventArgs) Handles btnFilesActive.Click
        Try
            Dim paths As List(Of String) = Globals.ThisAddIn.RunFilesPerDeptActive()
            ShowPerDeptResult(paths, "Отделы по файлам (активная)")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    ' === 4) ОТДЕЛЫ ПО ФАЙЛАМ — выбрать файл ===
    Private Sub btnFilesFile_Click(sender As Object, e As RibbonControlEventArgs) Handles btnFilesFile.Click
        Using dlg As New OpenFileDialog()
            dlg.Title = "Выберите файл Excel"
            dlg.Filter = "Excel книги (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
            dlg.Multiselect = False
            If dlg.ShowDialog() = DialogResult.OK Then
                Try
                    Dim paths As List(Of String) = Globals.ThisAddIn.RunFilesPerDeptFromFile(dlg.FileName)
                    ShowPerDeptResult(paths, "Отделы по файлам (файл)")
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                End Try
            End If
        End Using
    End Sub

    ' === ПРОСТАВИТЬ ОТПУСКА ===
    Private Sub btnApplyVacations_Click(sender As Object, e As RibbonControlEventArgs) Handles btnApplyVacations.Click
        Using dlg As New OpenFileDialog()
            dlg.Title = "Выберите файл с отпусками"
            dlg.Filter = "Excel книги (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
            dlg.Multiselect = False
            If dlg.ShowDialog() = DialogResult.OK Then
                Try
                    Dim filled As Integer = Globals.ThisAddIn.ApplyVacationsFromFile(dlg.FileName)
                    Dim caption As String = "Проставить отпуска"
                    Dim message As String
                    If filled > 0 Then
                        message = $"Готово: обновлено {filled} строк."
                    Else
                        message = "Совпадений не найдено."
                    End If
                    MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Проставить отпуска", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                End Try
            End If
        End Using
    End Sub

    ' === ПРОСТАВИТЬ ГРАФИКИ РАБОТЫ ===
    Private Sub btnApplyWorkSchedules_Click(sender As Object, e As RibbonControlEventArgs) Handles btnApplyWorkSchedules.Click
        Using dlg As New OpenFileDialog()
            dlg.Title = "Выберите файл с графиками работы"
            dlg.Filter = "Excel книги (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
            dlg.Multiselect = False
            If dlg.ShowDialog() = DialogResult.OK Then
                Try
                    Dim filled As Integer = Globals.ThisAddIn.ApplyWorkSchedulesFromFile(dlg.FileName)
                    Dim caption As String = "Проставить графики работы"
                    Dim message As String
                    If filled > 0 Then
                        message = $"Готово: обновлено {filled} строк."
                    Else
                        message = "Совпадений не найдено."
                    End If
                    MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Проставить графики работы", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                End Try
            End If
        End Using
    End Sub

    ' === Общее представление результата для режима «по файлам» ===
    Private Sub ShowPerDeptResult(paths As List(Of String), caption As String)
        If paths Is Nothing OrElse paths.Count = 0 Then
            MessageBox.Show("Не создано ни одного файла.", caption,
                            MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim folder As String = Path.GetDirectoryName(paths(0))
        Dim msg As String = "Созданы файлы:" & Environment.NewLine & String.Join(Environment.NewLine, paths)
        MessageBox.Show(msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Information)

        ' Откроем папку с результатами
        Try
            Process.Start("explorer.exe", folder)
        Catch
        End Try
    End Sub

End Class