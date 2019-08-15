## This code adds a Summary/Index Sheet to the workbook

Sub Index()

Dim ws As Worksheet, i As Integer

Worksheets.Add(before:=Worksheets(1)).Name = "Index"

For Each ws In ThisWorkbook.Worksheets
    If ws.Name <> "Index" Then
        i = i + 1
        Sheets("Index").Range("A" & i).Value = ws.Name
        Sheets("Index").Hyperlinks.Add Anchor:=Range("A" & i), Address:="", SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
    End If
    
Next ws

Sheets("Index").Columns("A").AutoFit

End Sub



## The code bellow adds an hyperlink to go back to active sheet. 
## Combined they can create an easy way to navigate on a big worksheet (Just be careful to not overwrite the index)

Sub CreateSummary()

Dim ws As Worksheet
Dim i As Integer
i = 0

For Each x In Worksheets
i = i + 1

If i = 1 Then GoTo Donothing
    With Worksheets(i)
        .Range("A1").Value = "Back to " & ActiveSheet.Name
        .Hyperlinks.Add Sheets(x.Name).Range("A1"), "", _
        "'" & ActiveSheet.Name & "'" & "!" & ActiveCell.Address, _
        ScreenTip:="Return to " & ActiveSheet.Name
    End With
    
    Donothing:
        Next x
    
End Sub
