Imports Excel = Microsoft.Office.Interop.Excel

Public Class ExcelController
    Public Function FindValueInExcel(ByVal sTextToFind)
        Dim currentFind As Excel.Range = Nothing
        Dim firstFind As Excel.Range = Nothing
        Dim xlappFirstFile As Excel.Application = Nothing
        Dim sSearchedValue As String

        xlappFirstFile = CreateObject("Excel.Application")

        xlappFirstFile.Workbooks.Open("D:\Sample.xlsx")

        Dim rngSearchValue As Excel.Range = xlappFirstFile.Range("A1", "C5")

        currentFind = rngSearchValue.Find(sTextToFind, ,
            Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False)

        If Not currentFind Is Nothing Then
            sSearchedValue = DirectCast(DirectCast(currentFind.EntireRow, Microsoft.Office.Interop.Excel.Range).Value2, System.Object)(1, 3).ToString()
        Else
            sSearchedValue = ""
        End If

        Return sSearchedValue
    End Function

    Public Sub SearchText()
        Dim oXL As Excel.Application
        Dim oWB As Excel.Workbook
        Dim oSheet As Excel.Worksheet

        Try
            Dim File_name As String = "C:\test.xls"
            oXL = CreateObject("Excel.Application")
            oWB = oXL.Workbooks.Open(File_name)
            oSheet = oWB.Worksheets(1)

            Dim oRng As Excel.Range = GetSpecifiedRange("get", oSheet)
            If oRng IsNot Nothing Then
                MessageBox.Show("Text found, position is Row-" & oRng.Row & " and column-" & oRng.Column)
            Else
                MessageBox.Show("Text is not found")
            End If
            oWB.Close()

            oSheet = Nothing
            oWB = Nothing
            oXL.Quit()
        Catch ex As Exception
            If oSheet IsNot Nothing Then
                oSheet = Nothing
            End If
            If oWB IsNot Nothing Then
                oWB = Nothing
            End If
            If oXL IsNot Nothing Then
                oXL.Quit()
            End If
        End Try
    End Sub

    Private Function GetSpecifiedRange(ByVal matchStr As String, ByVal objWs As Excel.Worksheet) As Excel.Range
        Dim currentFind As Excel.Range = Nothing
        Dim firstFind As Excel.Range = Nothing
        currentFind = objWs.Range("A1:AM100").Find(matchStr, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False)

        Return currentFind
    End Function
End Class
