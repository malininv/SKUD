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

        Dim msg As String = "Созданы файлы:" & Environment.NewLine & String.Join(Environment.NewLine, paths)
        MessageBox.Show(msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' Вспомогательная функция для поиска папки проекта
    Private Function FindProjectPath() As String
        ' Метод 1: Проверяем известный путь к проекту (для разработки)
        Dim knownPath As String = "C:\Users\Vladislav\source\repos\SKUD"
        If Directory.Exists(Path.Combine(knownPath, "src")) Then
            Return knownPath
        End If

        ' Метод 2: Ищем папку SKUD в стандартных местах
        Dim userProfile As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        Dim userName As String = Environment.UserName
        Dim possiblePaths As String() = {
            Path.Combine(userProfile, "source", "repos", "SKUD"),
            Path.Combine(userProfile, "Documents", "source", "repos", "SKUD"),
            Path.Combine("C:\", "Users", userName, "source", "repos", "SKUD"),
            Path.Combine("C:\", "Users", userName, "Documents", "source", "repos", "SKUD")
        }

        For Each pathItem In possiblePaths
            If Directory.Exists(pathItem) AndAlso Directory.Exists(Path.Combine(pathItem, "src")) Then
                Return pathItem
            End If
        Next

        ' Метод 3: Путь к сборке - поднимаемся по дереву
        Dim assemblyPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        If Not String.IsNullOrEmpty(assemblyPath) Then
            Dim searchPath As String = assemblyPath
            ' Поднимаемся по дереву папок, ищем папку src или файл .vbproj
            For i As Integer = 0 To 15
                If String.IsNullOrEmpty(searchPath) Then Exit For

                ' Проверяем наличие папки src
                If Directory.Exists(Path.Combine(searchPath, "src")) Then
                    Return searchPath
                End If

                ' Ищем файл .vbproj или .sln
                Dim vbprojFiles As String() = Directory.GetFiles(searchPath, "*.vbproj")
                Dim slnFiles As String() = Directory.GetFiles(searchPath, "*.sln")
                If (vbprojFiles.Length > 0 OrElse slnFiles.Length > 0) AndAlso Directory.Exists(Path.Combine(searchPath, "src")) Then
                    Return searchPath
                End If

                Dim parentDir As DirectoryInfo = Directory.GetParent(searchPath)
                If parentDir Is Nothing Then Exit For
                searchPath = parentDir.FullName
            Next
        End If

        ' Метод 4: Пробуем через текущую рабочую директорию
        Dim currentDir As String = Directory.GetCurrentDirectory()
        Dim searchFromCurrent As String = currentDir
        For i As Integer = 0 To 10
            If Directory.Exists(Path.Combine(searchFromCurrent, "src")) Then
                Return searchFromCurrent
            End If
            Dim parentDir As DirectoryInfo = Directory.GetParent(searchFromCurrent)
            If parentDir Is Nothing Then Exit For
            searchFromCurrent = parentDir.FullName
        Next

        ' Метод 5: Пробуем через AppDomain
        Try
            Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
            If Not String.IsNullOrEmpty(baseDir) Then
                Dim searchFromBase As String = baseDir
                For i As Integer = 0 To 10
                    If Directory.Exists(Path.Combine(searchFromBase, "src")) Then
                        Return searchFromBase
                    End If
                    Dim parentDir As DirectoryInfo = Directory.GetParent(searchFromBase)
                    If parentDir Is Nothing Then Exit For
                    searchFromBase = parentDir.FullName
                Next
            End If
        Catch
        End Try

        Return Nothing
    End Function

    ' Вспомогательная функция для сохранения скриншота из ресурсов
    Private Sub SaveScreenshotFromResources(tempDir As String, resourceName As String, fileName As String)
        Dim saved As Boolean = False

        ' Метод 1: Пробуем через рефлексию - прямое обращение к свойствам My.Resources
        Try
            Dim resourceType As Type = GetType(My.Resources.Resources)
            Dim prop As System.Reflection.PropertyInfo = resourceType.GetProperty(resourceName, System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Static)
            If prop IsNot Nothing Then
                Dim resourceObject As Object = prop.GetValue(Nothing, Nothing)
                If resourceObject IsNot Nothing AndAlso TypeOf resourceObject Is System.Drawing.Bitmap Then
                    Dim screenshotImage As System.Drawing.Bitmap = DirectCast(resourceObject, System.Drawing.Bitmap)
                    Dim tempScreenshotPath As String = Path.Combine(tempDir, fileName)
                    screenshotImage.Save(tempScreenshotPath, System.Drawing.Imaging.ImageFormat.Png)
                    saved = True
                End If
            End If
        Catch
        End Try

        ' Метод 2: Если первый не сработал, пробуем через ResourceManager
        If Not saved Then
            Try
                Dim resourceManager As System.Resources.ResourceManager = My.Resources.ResourceManager
                Dim resourceObject As Object = resourceManager.GetObject(resourceName)

                If resourceObject IsNot Nothing AndAlso TypeOf resourceObject Is System.Drawing.Bitmap Then
                    Dim screenshotImage As System.Drawing.Bitmap = DirectCast(resourceObject, System.Drawing.Bitmap)
                    Dim tempScreenshotPath As String = Path.Combine(tempDir, fileName)
                    screenshotImage.Save(tempScreenshotPath, System.Drawing.Imaging.ImageFormat.Png)
                    saved = True
                End If
            Catch
            End Try
        End If

        ' Метод 3: Если ресурсы не загрузились, пробуем скопировать из файла (fallback)
        If Not saved Then
            Try
                Dim projectPath As String = FindProjectPath()
                If Not String.IsNullOrEmpty(projectPath) Then
                    Dim screenshotPath As String = Path.Combine(projectPath, "src", fileName)
                    If File.Exists(screenshotPath) Then
                        Dim tempScreenshotPath As String = Path.Combine(tempDir, fileName)
                        File.Copy(screenshotPath, tempScreenshotPath, True)
                        saved = True
                    End If
                End If
            Catch
                ' Игнорируем ошибки при копировании из файла
            End Try
        End If
    End Sub

    ' === ПОКАЗАТЬ ИНСТРУКЦИЮ ПОЛЬЗОВАТЕЛЯ ===
    Private Sub btnShowInstructions_Click(sender As Object, e As RibbonControlEventArgs) Handles btnShowInstructions.Click
        Try
            Dim screenshotFileName As String = "Excel_panel.png"

            ' Создаем временную папку
            Dim tempDir As String = Path.Combine(Path.GetTempPath(), "SKUD_Instructions")
            If Not Directory.Exists(tempDir) Then
                Directory.CreateDirectory(tempDir)
            End If

            ' Получаем HTML из ресурсов или из файла
            Dim htmlContent As String = Nothing
            Try
                ' Пробуем через рефлексию
                Dim resourceType As Type = GetType(My.Resources.Resources)
                Dim prop As System.Reflection.PropertyInfo = resourceType.GetProperty("user_instruction", System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Static)
                If prop IsNot Nothing Then
                    htmlContent = DirectCast(prop.GetValue(Nothing, Nothing), String)
                End If
            Catch
            End Try
            
            ' Если через рефлексию не получилось, пробуем напрямую
            If String.IsNullOrEmpty(htmlContent) Then
                Try
                    htmlContent = My.Resources.user_instruction
                Catch
                End Try
            End If
            
            ' Проверяем кодировку
            If htmlContent IsNot Nothing AndAlso htmlContent.Contains("Рћ") Then
                ' Похоже на неправильную кодировку, пробуем прочитать из файла
                htmlContent = Nothing
            End If

            ' Если ресурс не найден или имеет неправильную кодировку, пробуем прочитать из файла
            If String.IsNullOrEmpty(htmlContent) Then
                Dim projectPath As String = FindProjectPath()
                If Not String.IsNullOrEmpty(projectPath) Then
                    Dim htmlPath As String = Path.Combine(projectPath, "src", "user_instruction.html")
                    If File.Exists(htmlPath) Then
                        ' Читаем с автоматическим определением кодировки
                        htmlContent = File.ReadAllText(htmlPath, System.Text.Encoding.UTF8)
                    End If
                End If
            End If

            If String.IsNullOrEmpty(htmlContent) Then
                Dim errorMsg As String = "Не удалось загрузить инструкцию из ресурсов или файла." & Environment.NewLine & Environment.NewLine
                errorMsg &= "Проверьте:" & Environment.NewLine
                errorMsg &= "1. Файл user_instruction.html добавлен в ресурсы проекта" & Environment.NewLine
                errorMsg &= "2. Файл находится в папке src проекта"
                MessageBox.Show(errorMsg, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Сохраняем HTML во временный файл с правильной кодировкой UTF-8 с BOM
            Dim tempHtmlPath As String = Path.Combine(tempDir, "SKUD_UserInstruction.html")
            Using writer As New System.IO.StreamWriter(tempHtmlPath, False, New System.Text.UTF8Encoding(True))
                writer.Write(htmlContent)
            End Using

            ' Сохраняем все изображения из ресурсов во временную папку
            SaveScreenshotFromResources(tempDir, "Excel_panel", "Excel_panel.png")
            SaveScreenshotFromResources(tempDir, "URV", "URV.png")
            SaveScreenshotFromResources(tempDir, "Nastroiki", "Nastroiki.png")
            SaveScreenshotFromResources(tempDir, "Sotrudniki", "Sotrudniki.png")
            SaveScreenshotFromResources(tempDir, "Export_v_Excel", "Export_v_Excel.png")
            SaveScreenshotFromResources(tempDir, "Export_parametri", "Export_parametri.png")

            ' Открываем в браузере по умолчанию
            Process.Start(tempHtmlPath)

        Catch ex As Exception
            MessageBox.Show($"Ошибка при открытии инструкции: {ex.Message}", "Ошибка",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' === ПОКАЗАТЬ ИНСТРУКЦИЮ ПО ВЫГРУЗКЕ ОТЧЕТА ИЗ СКУД ===
    Private Sub btnShowExportReport_Click(sender As Object, e As RibbonControlEventArgs) Handles btnShowExportReport.Click
        Try
            ' Создаем временную папку
            Dim tempDir As String = Path.Combine(Path.GetTempPath(), "SKUD_Instructions")
            If Not Directory.Exists(tempDir) Then
                Directory.CreateDirectory(tempDir)
            End If

            ' Получаем HTML из ресурсов или из файла
            Dim htmlContent As String = Nothing
            Try
                ' Пробуем загрузить из ресурсов через ResourceManager
                Dim resourceManager As System.Resources.ResourceManager = My.Resources.ResourceManager
                Dim resourceObject As Object = resourceManager.GetObject("export_report_instruction")
                If resourceObject IsNot Nothing AndAlso TypeOf resourceObject Is String Then
                    htmlContent = DirectCast(resourceObject, String)
                    ' Если строка содержит неправильную кодировку, пробуем перекодировать
                    If htmlContent IsNot Nothing AndAlso htmlContent.Contains("Рћ") Then
                        ' Похоже на неправильную кодировку, пробуем прочитать из файла
                        htmlContent = Nothing
                    End If
                End If
            Catch
                ' Ресурс не найден, пробуем прочитать из файла
            End Try

            ' Если ресурс не найден или имеет неправильную кодировку, пробуем прочитать из файла
            If String.IsNullOrEmpty(htmlContent) Then
                Dim projectPath As String = FindProjectPath()
                If Not String.IsNullOrEmpty(projectPath) Then
                    Dim htmlPath As String = Path.Combine(projectPath, "src", "export_report_instruction.html")
                    If File.Exists(htmlPath) Then
                        ' Читаем с автоматическим определением кодировки
                        htmlContent = File.ReadAllText(htmlPath, System.Text.Encoding.UTF8)
                    End If
                End If
            End If

            If String.IsNullOrEmpty(htmlContent) Then
                Dim errorMsg As String = "Не удалось загрузить инструкцию из ресурсов или файла." & Environment.NewLine & Environment.NewLine
                errorMsg &= "Проверьте:" & Environment.NewLine
                errorMsg &= "1. Файл export_report_instruction.html добавлен в ресурсы проекта" & Environment.NewLine
                errorMsg &= "2. Файл находится в папке src проекта"
                MessageBox.Show(errorMsg, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Сохраняем HTML во временный файл с правильной кодировкой UTF-8 с BOM
            Dim tempHtmlPath As String = Path.Combine(tempDir, "SKUD_ExportReportInstruction.html")
            Using writer As New System.IO.StreamWriter(tempHtmlPath, False, New System.Text.UTF8Encoding(True))
                writer.Write(htmlContent)
            End Using

            ' Сохраняем все изображения из ресурсов во временную папку
            SaveScreenshotFromResources(tempDir, "URV", "URV.png")
            SaveScreenshotFromResources(tempDir, "Ustanovit_soed", "Ustanovit_soed.png")
            SaveScreenshotFromResources(tempDir, "Nastroiki", "Nastroiki.png")
            SaveScreenshotFromResources(tempDir, "Sotrudniki", "Sotrudniki.png")
            SaveScreenshotFromResources(tempDir, "Export_v_Excel", "Export_v_Excel.png")
            SaveScreenshotFromResources(tempDir, "Export_parametri", "Export_parametri.png")

            ' Открываем в браузере по умолчанию
            Process.Start(tempHtmlPath)

        Catch ex As Exception
            MessageBox.Show($"Ошибка при открытии инструкции: {ex.Message}", "Ошибка",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' === ПОКАЗАТЬ ИНСТРУКЦИЮ ПО ВЫГРУЗКЕ ОТПУСКОВ И ГРАФИКОВ ===
    Private Sub btnShowExportInstructions_Click(sender As Object, e As RibbonControlEventArgs) Handles btnShowExportInstructions.Click
        Try
            ' Создаем временную папку
            Dim tempDir As String = Path.Combine(Path.GetTempPath(), "SKUD_Instructions")
            If Not Directory.Exists(tempDir) Then
                Directory.CreateDirectory(tempDir)
            End If

            ' Получаем HTML из ресурсов или из файла
            Dim htmlContent As String = Nothing
            Try
                ' Пробуем через рефлексию
                Dim resourceType As Type = GetType(My.Resources.Resources)
                Dim prop As System.Reflection.PropertyInfo = resourceType.GetProperty("export_instruction", System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Static)
                If prop IsNot Nothing Then
                    htmlContent = DirectCast(prop.GetValue(Nothing, Nothing), String)
                End If
            Catch
            End Try
            
            ' Если через рефлексию не получилось, пробуем напрямую
            If String.IsNullOrEmpty(htmlContent) Then
                Try
                    htmlContent = My.Resources.export_instruction
                Catch
                End Try
            End If
            
            ' Проверяем кодировку
            If htmlContent IsNot Nothing AndAlso htmlContent.Contains("Рћ") Then
                ' Похоже на неправильную кодировку, пробуем прочитать из файла
                htmlContent = Nothing
            End If

            ' Если ресурс не найден или имеет неправильную кодировку, пробуем прочитать из файла
            If String.IsNullOrEmpty(htmlContent) Then
                Dim projectPath As String = FindProjectPath()
                If Not String.IsNullOrEmpty(projectPath) Then
                    Dim htmlPath As String = Path.Combine(projectPath, "src", "export_instruction.html")
                    If File.Exists(htmlPath) Then
                        ' Читаем с автоматическим определением кодировки
                        htmlContent = File.ReadAllText(htmlPath, System.Text.Encoding.UTF8)
                    End If
                End If
            End If

            If String.IsNullOrEmpty(htmlContent) Then
                Dim errorMsg As String = "Не удалось загрузить инструкцию из ресурсов или файла." & Environment.NewLine & Environment.NewLine
                errorMsg &= "Проверьте:" & Environment.NewLine
                errorMsg &= "1. Файл export_instruction.html добавлен в ресурсы проекта" & Environment.NewLine
                errorMsg &= "2. Файл находится в папке src проекта"
                MessageBox.Show(errorMsg, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Сохраняем HTML во временный файл с правильной кодировкой UTF-8 с BOM
            Dim tempHtmlPath As String = Path.Combine(tempDir, "SKUD_ExportInstruction.html")
            Using writer As New System.IO.StreamWriter(tempHtmlPath, False, New System.Text.UTF8Encoding(True))
                writer.Write(htmlContent)
            End Using

            ' Сохраняем изображения из ресурсов во временную папку
            SaveScreenshotFromResources(tempDir, "Otpusk", "Otpusk.png")
            SaveScreenshotFromResources(tempDir, "Generator", "Generator.png")
            SaveScreenshotFromResources(tempDir, "Grafik_raboti_1", "Grafik_raboti_1.png")
            SaveScreenshotFromResources(tempDir, "Grafik_raboti_2", "Grafik_raboti_2.png")

            ' Открываем в браузере по умолчанию
            Process.Start(tempHtmlPath)

        Catch ex As Exception
            MessageBox.Show($"Ошибка при открытии инструкции: {ex.Message}", "Ошибка",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class