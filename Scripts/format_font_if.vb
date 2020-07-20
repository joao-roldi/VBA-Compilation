Private Sub Format_range()
    'Source: https://excel.tips.net/T002380_Changing_Font_Face_and_Size_Conditionally.html

    Dim rng As Range
    Dim rCell As Range

    Set rng = Range("F3:L10")

    For Each rCell In rng
        If Len(rCell.Text) > 2 And _
          rCell.Value = "N/A" Then
            rCell.Font.Name = "Calibri"
            rCell.Font.Size = 11
            rCell.Font.Color = RGB(0, 0, 0)
        ElseIf rCell.Value = "P" Then
            rCell.Font.Name = "Wingdings 2"
            rCell.Font.Bold = True
            rCell.Font.Size = 11
            rCell.Font.Color = RGB(49, 155, 66)
        ElseIf rCell.Value = "O" Then
            rCell.Font.Name = "Wingdings 2"
            rCell.Font.Bold = True
            rCell.Font.Size = 11
            rCell.Font.Color = 255
        Else
            rCell.Font.Name = "Calibri"
            rCell.Font.Size = 11
        End If
    Next
End Sub