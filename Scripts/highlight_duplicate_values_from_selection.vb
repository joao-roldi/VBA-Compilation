Sub duplicate_values_from_selection()
'This is useful when you have a lot of rows and the conditional format is impracticable

Dim myRange As Range
Dim i As Integer
Dim j As Integer
Dim myCell As Range
Set myRange = Selection

For Each myCell In myRange
  If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
    myCell.Interior.ColorIndex = 3
  End If
Next
End Sub
