Sub remove_fake_empty()
' This code was found on: https://stackoverflow.com/questions/15984580/excel-telling-me-my-blank-cells-arent-blank

With Selection
    Set c = .Find("", LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then
        firstAddress = c.Address
        Do
            c.Value = ""
            Set c = .FindNext(c)
            If c Is Nothing Then Exit Do
        Loop While c.Address <> firstAddress
    End If
End With

MsgBox "Complete!"
End Sub