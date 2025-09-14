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
        Me.btnFilesActive = Me.Factory.CreateRibbonButton
        Me.btnFilesFile = Me.Factory.CreateRibbonButton
        Me.btnSheetsFile = Me.Factory.CreateRibbonButton
        Me.btnSheetsActive = Me.Factory.CreateRibbonButton
        Me.Box1 = Me.Factory.CreateRibbonBox
        Me.Box2 = Me.Factory.CreateRibbonBox
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Box1.SuspendLayout()
        Me.Box2.SuspendLayout()
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
        Me.Group1.Items.Add(Me.Box1)
        Me.Group1.Items.Add(Me.Box2)
        Me.Group1.Label = "Отчет по проходам"
        Me.Group1.Name = "Group1"
        '
        'btnFilesActive
        '
        Me.btnFilesActive.Label = "По файлам (эта книга)"
        Me.btnFilesActive.Name = "btnFilesActive"
        Me.btnFilesActive.OfficeImageId = "Folder"
        Me.btnFilesActive.ShowImage = True
        '
        'btnFilesFile
        '
        Me.btnFilesFile.Label = "По файлам (выбрать книгу)"
        Me.btnFilesFile.Name = "btnFilesFile"
        Me.btnFilesFile.OfficeImageId = "FileOpen"
        Me.btnFilesFile.ShowImage = True
        '
        'btnSheetsFile
        '
        Me.btnSheetsFile.Label = "По листам (выбрать книгу)"
        Me.btnSheetsFile.Name = "btnSheetsFile"
        Me.btnSheetsFile.OfficeImageId = "FileOpen"
        Me.btnSheetsFile.ShowImage = True
        '
        'btnSheetsActive
        '
        Me.btnSheetsActive.Label = "По листам (эта книга)"
        Me.btnSheetsActive.Name = "btnSheetsActive"
        Me.btnSheetsActive.OfficeImageId = "TableInsert"
        Me.btnSheetsActive.ShowImage = True
        '
        'Box1
        '
        Me.Box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box1.Items.Add(Me.btnSheetsActive)
        Me.Box1.Items.Add(Me.btnSheetsFile)
        Me.Box1.Name = "Box1"
        '
        'Box2
        '
        Me.Box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box2.Items.Add(Me.btnFilesFile)
        Me.Box2.Items.Add(Me.btnFilesActive)
        Me.Box2.Name = "Box2"
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
        Me.Box1.ResumeLayout(False)
        Me.Box1.PerformLayout()
        Me.Box2.ResumeLayout(False)
        Me.Box2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnFilesActive As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSheetsActive As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSheetsFile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnFilesFile As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Box1 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents Box2 As Microsoft.Office.Tools.Ribbon.RibbonBox
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
