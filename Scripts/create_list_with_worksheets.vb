Sub ListSheetNamesInNewWorkbook()
'Found on: https://www.datanumen.com/blogs/3-quick-ways-to-get-a-list-of-all-worksheet-names-in-an-excel-workbook/
'You have to add this script in the "ThisWorkbook" module for it to work properly

    Dim objNewWorkbook As Workbook
    Dim objNewWorksheet As Worksheet

    Set objNewWorkbook = Excel.Application.Workbooks.Add
    Set objNewWorksheet = objNewWorkbook.Sheets(1)

    For i = 1 To ThisWorkbook.Sheets.Count
        objNewWorksheet.Cells(i, 1) = i
        objNewWorksheet.Cells(i, 2) = ThisWorkbook.Sheets(i).Name
    Next i

    With objNewWorksheet
         .Rows(1).Insert
         .Cells(1, 1) = "INDEX"
         .Cells(1, 1).Font.Bold = True
         .Cells(1, 2) = "NAME"
         .Cells(1, 2).Font.Bold = True
         .Columns("A:B").AutoFit
    End With
End Sub