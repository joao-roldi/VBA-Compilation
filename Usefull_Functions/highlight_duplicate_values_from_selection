Sub duplicate_values_from_selection()

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
