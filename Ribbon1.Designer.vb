Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.btnSheetsActive = Me.Factory.CreateRibbonButton
        Me.btnSheetsFile = Me.Factory.CreateRibbonButton
        Me.btnFilesActive = Me.Factory.CreateRibbonButton
        Me.btnFilesFile = Me.Factory.CreateRibbonButton
        Me.ButtonGroup1 = Me.Factory.CreateRibbonButtonGroup
        Me.ButtonGroup2 = Me.Factory.CreateRibbonButtonGroup
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.ButtonGroup1.SuspendLayout()
        Me.ButtonGroup2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "СКУД"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.ButtonGroup1)
        Me.Group1.Items.Add(Me.ButtonGroup2)
        Me.Group1.Label = "Отчет по проходам"
        Me.Group1.Name = "Group1"
        '
        'btnSheetsActive
        '
        Me.btnSheetsActive.Label = "По листам (это книга)"
        Me.btnSheetsActive.Name = "btnSheetsActive"
        '
        'btnSheetsFile
        '
        Me.btnSheetsFile.Label = "По листам (выбрать книгу)"
        Me.btnSheetsFile.Name = "btnSheetsFile"
        '
        'btnFilesActive
        '
        Me.btnFilesActive.Label = "По файлам (эта книга)"
        Me.btnFilesActive.Name = "btnFilesActive"
        '
        'btnFilesFile
        '
        Me.btnFilesFile.Label = "По файлам (выбрать книгу)"
        Me.btnFilesFile.Name = "btnFilesFile"
        '
        'ButtonGroup1
        '
        Me.ButtonGroup1.Items.Add(Me.btnSheetsActive)
        Me.ButtonGroup1.Items.Add(Me.btnSheetsFile)
        Me.ButtonGroup1.Name = "ButtonGroup1"
        '
        'ButtonGroup2
        '
        Me.ButtonGroup2.Items.Add(Me.btnFilesActive)
        Me.ButtonGroup2.Items.Add(Me.btnFilesFile)
        Me.ButtonGroup2.Name = "ButtonGroup2"
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
        Me.ButtonGroup1.ResumeLayout(False)
        Me.ButtonGroup1.PerformLayout()
        Me.ButtonGroup2.ResumeLayout(False)
        Me.ButtonGroup2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnSheetsActive As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSheetsFile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnFilesActive As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonGroup1 As Microsoft.Office.Tools.Ribbon.RibbonButtonGroup
    Friend WithEvents ButtonGroup2 As Microsoft.Office.Tools.Ribbon.RibbonButtonGroup
    Friend WithEvents btnFilesFile As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
