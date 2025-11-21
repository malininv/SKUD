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

    ' === ПОКАЗАТЬ ИНСТРУКЦИЮ ===
    Private Sub btnShowInstructions_Click(sender As Object, e As RibbonControlEventArgs) Handles btnShowInstructions.Click
        Try
            ' Создаем HTML-страницу с инструкцией
            Dim instructionHtml As String = CreateInstructionHtml()
            
            ' Создаем временный файл
            Dim tempPath As String = Path.Combine(Path.GetTempPath(), "SKUD_Instruction.html")
            File.WriteAllText(tempPath, instructionHtml, System.Text.Encoding.UTF8)
            
            ' Открываем в браузере по умолчанию
            Process.Start(tempPath)
            
        Catch ex As Exception
            MessageBox.Show($"Ошибка при открытии инструкции: {ex.Message}", "Ошибка", 
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Создает HTML-страницу с инструкцией
    Private Function CreateInstructionHtml() As String
        Return "<!DOCTYPE html>
<html lang=""ru"">
<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>Инструкция по использованию СКУД</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }
        h1 { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
        h2 { color: #34495e; margin-top: 30px; }
        h3 { color: #7f8c8d; }
        .step { background-color: #f8f9fa; padding: 15px; margin: 10px 0; border-left: 4px solid #3498db; }
        .warning { background-color: #fff3cd; padding: 15px; margin: 10px 0; border-left: 4px solid #ffc107; }
        .success { background-color: #d4edda; padding: 15px; margin: 10px 0; border-left: 4px solid #28a745; }
        .code { background-color: #f1f2f6; padding: 10px; font-family: monospace; border-radius: 4px; }
        ul { padding-left: 20px; }
        li { margin: 5px 0; }
        .image-placeholder { 
            background-color: #e9ecef; 
            border: 2px dashed #6c757d; 
            padding: 40px; 
            text-align: center; 
            margin: 20px 0;
            border-radius: 8px;
        }
    </style>
</head>
<body>
    <h1>📋 Инструкция по использованию СКУД</h1>
    
    <h2>🎯 Общее описание</h2>
    <p>Надстройка СКУД предназначена для автоматической обработки данных системы контроля и управления доступом, 
    создания отчетов по отделам и проставления отпусков и графиков работы.</p>
    
    <h2>🔧 Основные функции</h2>
    
    <h3>1. Отделы по листам</h3>
    <div class=""step"">
        <strong>Назначение:</strong> Создает один файл Excel с отдельными листами для каждого отдела.<br>
        <strong>Использование:</strong>
        <ul>
            <li><strong>Активная книга:</strong> Обрабатывает текущую открытую книгу Excel</li>
            <li><strong>Выбрать файл:</strong> Позволяет выбрать файл для обработки</li>
        </ul>
    </div>
    
    <h3>2. Отделы по файлам</h3>
    <div class=""step"">
        <strong>Назначение:</strong> Создает отдельный файл Excel для каждого отдела в папке.<br>
        <strong>Использование:</strong>
        <ul>
            <li><strong>Активная книга:</strong> Обрабатывает текущую открытую книгу Excel</li>
            <li><strong>Выбрать файл:</strong> Позволяет выбрать файл для обработки</li>
        </ul>
    </div>
    
    <h3>3. Проставление отпусков</h3>
    <div class=""step"">
        <strong>Назначение:</strong> Автоматически проставляет причины отсутствия сотрудников на основе файла отпусков.<br>
        <strong>Формат файла отпусков:</strong>
        <ul>
            <li>Колонка A: ФИО сотрудника</li>
            <li>Колонка E: Дата начала отпуска</li>
            <li>Колонка F: Дата окончания отпуска</li>
            <li>Колонка G: Причина отсутствия</li>
        </ul>
    </div>
    
    <h3>4. Проставление графиков работы</h3>
    <div class=""step"">
        <strong>Назначение:</strong> Автоматически проставляет графики работы сотрудников.<br>
        <strong>Формат файла графиков:</strong>
        <ul>
            <li>Колонка C: ФИО сотрудника (начиная со строки 7)</li>
            <li>Колонка F: График работы (начиная со строки 7)</li>
        </ul>
    </div>
    
    <h2>📊 Структура данных</h2>
    <div class=""warning"">
        <strong>Важно!</strong> Исходный файл выгрузки из СКУД должен содержать следующие колонки:
        <ul>
            <li>1. Фирма</li>
            <li>2. Подразделение</li>
            <li>3. Сотрудник</li>
            <li>4. Должность</li>
            <li>5. Таб.№</li>
            <li>6. Дата</li>
            <li>7. Находился в здании</li>
            <li>8. Прогулял</li>
            <li>9. Причины не выхода</li>
            <li>10. Комм. причины отсутствия</li>
            <li>11. Начало дня</li>
            <li>12. Конец дня</li>
            <li>13. Работа в праздничные дни</li>
            <li>14. Фактическая переработка</li>
        </ul>
    </div>
    
    <h2>⚙️ Настройки в СКУД</h2>
    <div class=""step"">
        <strong>Важно!</strong> Для корректной работы надстройки в системе СКУД должны быть настроены следующие параметры:
    </div>
    
    <div class=""image-placeholder"">
        📷 <strong>Скриншот 1: Параметры в Учетре рабочего времени в СКУД</strong><br>
        Здесь будет изображение с настройками системы СКУД для корректной выгрузки данных
    </div>
    
    <div class=""image-placeholder"">
        📷 <strong>Скриншот 2: Настройки экспорта отчета в Excel в СКУД. Выбраны все колонки.</strong><br>
        Здесь будет изображение с примером правильной структуры файла выгрузки
    </div>

    <h2>🎨 Особенности обработки</h2>
    
    <h3>Автоматическое форматирование</h3>
    <ul>
        <li>Удаление выходных дней с нулевым временем</li>
        <li>Подсветка ячеек с нулевым временем (желтый фон, красный шрифт)</li>
        <li>Цветовая индикация фактической переработки:
            <ul>
                <li>🟢 Зеленый: положительная переработка</li>
                <li>🟡 Желтый: небольшая недоработка (-1 до 0 часов)</li>
                <li>🔴 Красный: значительная недоработка (более -1 часа)</li>
            </ul>
        </li>
    </ul>
    
    <h3>Создание листов ""_нет_прохода""</h3>
    <div class=""success"">
        Система автоматически создает отдельные листы с суффиксом ""_нет_прохода"" для отделов, 
        где сотрудники не проходили через систему контроля доступа за указанный период.
    </div>
    
    
    <h2>🆘 Поддержка</h2>
    <p>При возникновении проблем или вопросов обращайтесь к разработчику.</p>
    <p>Разработчик: Малинин Владислав</p>
    <p>e-mail: vladmalinin93@gmail.com</p>
    <p>Телефон: +7 (951) 187-67-10</p>
    
    <hr>
    <p><em>Версия: 1.0.0.9 | Дата обновления: " & DateTime.Now.ToString("dd.MM.yyyy") & "</em></p>
</body>
</html>"
    End Function

End Class