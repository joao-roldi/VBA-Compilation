' For values

Sub format_cell_if_number()
Dim rng As Range
For Each rng In Selection

    If rng.Value <> "" Then
        If IsNumeric(rng.Value) Then
            If Abs(rng.Value) > (0.15) Then
                rng.Font.Color = vbRed
                rng.Font.Bold = True
            Else
                rng.Font.Color = vbBlack
                rng.Font.Bold = False
            End If
        End If
    End If
    
Next rng
End Sub


' For Text

Sub format_cell_if_text()
Dim rng As Range
For Each rng In Selection

    If rng.Value <> "" Then
            If rng.Text = "Não" Then
                rng.Font.Color = vbRed
                rng.Font.Bold = True

            End If
    End If
    
Next rng
End Sub
