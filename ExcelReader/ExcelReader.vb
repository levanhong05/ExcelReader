Imports Excel = Microsoft.Office.Interop.Excel

Public Class ExcelReader
    Private Sub btnSourceBrowse_Click(sender As Object, e As EventArgs) Handles btnSourceBrowse.Click
        Dim dialog As New FolderBrowserDialog()
        dialog.RootFolder = Environment.SpecialFolder.Desktop
        dialog.SelectedPath = "C:\"
        dialog.Description = "Select Excel Files Path"

        If (dialog.ShowDialog() = DialogResult.OK) Then
            txtSourcePath.Text = dialog.SelectedPath
        End If
    End Sub

    Private Sub btnDestinationBrowse_Click(sender As Object, e As EventArgs) Handles btnDestinationBrowse.Click
        Dim dialog As OpenFileDialog = New OpenFileDialog()
        dialog.Title = "Select BILAN Excel File Path"
        dialog.InitialDirectory = "C:\"
        dialog.Filter = "Excel 2007 files (*.xlsx)|*.xlsx|Excel 2003 files (*.xls)|*.xls"
        dialog.FilterIndex = 2
        dialog.RestoreDirectory = True

        If (dialog.ShowDialog() = DialogResult.OK) Then
            txtDestinationPath.Text = dialog.FileName
        End If
    End Sub

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        If (txtSourcePath.Text = "") Then
            MessageBox.Show("Please select the source path")
            txtSourcePath.Focus()
            Return
        End If

        If (txtDestinationPath.Text = "") Then
            MessageBox.Show("Please select the destination path")
            txtDestinationPath.Focus()
            Return
        End If

        Dim excelApp As New Excel.Application
        Dim desWorksheet As Excel.Worksheet
        Dim desWworkbook As Excel.Workbook

        desWworkbook = excelApp.Workbooks.Open(txtDestinationPath.Text)
        desWorksheet = desWworkbook.Worksheets("Feuil1")

        desWworkbook.Save()
        desWworkbook.Close()
        excelApp.Quit()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class
