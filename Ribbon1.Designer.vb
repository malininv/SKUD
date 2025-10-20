Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.btnSheetsActive = Me.Factory.CreateRibbonButton
        Me.btnSheetsFile = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.btnFilesActive = Me.Factory.CreateRibbonButton
        Me.btnFilesFile = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.btnApplyVacations = Me.Factory.CreateRibbonButton
        Me.btnApplyWorkSchedules = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.btnShowInstructions = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Label = "СКУД"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.btnSheetsActive)
        Me.Group1.Items.Add(Me.btnSheetsFile)
        Me.Group1.Label = "Отделы по листам"
        Me.Group1.Name = "Group1"
        '
        'btnSheetsActive
        '
        Me.btnSheetsActive.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnSheetsActive.Label = "Активная книга"
        Me.btnSheetsActive.Name = "btnSheetsActive"
        Me.btnSheetsActive.OfficeImageId = "TableInsert"
        Me.btnSheetsActive.ShowImage = True
        '
        'btnSheetsFile
        '
        Me.btnSheetsFile.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnSheetsFile.Label = "Выбрать файл"
        Me.btnSheetsFile.Name = "btnSheetsFile"
        Me.btnSheetsFile.OfficeImageId = "FileOpen"
        Me.btnSheetsFile.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btnFilesActive)
        Me.Group2.Items.Add(Me.btnFilesFile)
        Me.Group2.Label = "Отделы по файлам"
        Me.Group2.Name = "Group2"
        '
        'btnFilesActive
        '
        Me.btnFilesActive.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnFilesActive.Label = "Активная книга"
        Me.btnFilesActive.Name = "btnFilesActive"
        Me.btnFilesActive.OfficeImageId = "Folder"
        Me.btnFilesActive.ShowImage = True
        '
        'btnFilesFile
        '
        Me.btnFilesFile.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnFilesFile.Label = "Выбрать файл"
        Me.btnFilesFile.Name = "btnFilesFile"
        Me.btnFilesFile.OfficeImageId = "FileOpen"
        Me.btnFilesFile.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.btnApplyVacations)
        Me.Group3.Items.Add(Me.btnApplyWorkSchedules)
        Me.Group3.Label = "Отпуска"
        Me.Group3.Name = "Group3"
        '
        'btnApplyVacations
        '
        Me.btnApplyVacations.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnApplyVacations.Label = "Проставить отпуска"
        Me.btnApplyVacations.Name = "btnApplyVacations"
        Me.btnApplyVacations.OfficeImageId = "ColumnsDialog"
        Me.btnApplyVacations.ShowImage = True
        '
        'btnApplyWorkSchedules
        '
        Me.btnApplyWorkSchedules.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnApplyWorkSchedules.Label = "Графики работы"
        Me.btnApplyWorkSchedules.Name = "btnApplyWorkSchedules"
        Me.btnApplyWorkSchedules.OfficeImageId = "ColumnsDialog"
        Me.btnApplyWorkSchedules.ShowImage = True
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.btnShowInstructions)
        Me.Group4.Label = "Справка"
        Me.Group4.Name = "Group4"
        '
        'btnShowInstructions
        '
        Me.btnShowInstructions.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnShowInstructions.Label = "Инструкция"
        Me.btnShowInstructions.Name = "btnShowInstructions"
        Me.btnShowInstructions.OfficeImageId = "Help"
        Me.btnShowInstructions.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnFilesActive As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSheetsActive As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSheetsFile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnFilesFile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnApplyVacations As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnApplyWorkSchedules As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnShowInstructions As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
