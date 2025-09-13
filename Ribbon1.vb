Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub BtnFromActive_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnFromActive.Click
        Globals.ThisAddIn.RunReportForActiveWorkbook()
    End Sub

    Private Sub BtnFromFile_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnFromFile.Click
        Globals.ThisAddIn.RunReportForFile()
    End Sub
End Class
