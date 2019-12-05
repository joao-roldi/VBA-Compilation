Sub ExportSheetsToCSV()
    Dim xWs As Worksheet
    Dim xcsvFile As String
    For Each xWs In Application.ActiveWorkbook.Worksheets
        xWs.Copy
        xcsvFile = CurDir & "\" & xWs.Name & ".csv"
        Application.ActiveWorkbook.SaveAs Filename: = xcsvFile, _
        FileFormat: = xlCSV, CreateBackup: = False
        Application.ActiveWorkbook.Saved = True
        Application.ActiveWorkbook.Close
    Next
End Sub