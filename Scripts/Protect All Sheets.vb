Sub ProtectAll()

    Dim wSheet As Worksheet
    Dim Pwd As String

    Pwd = InputBox("Enter your password to protect all worksheets", "Password Input")
    For Each wSheet In Worksheets
        wSheet.Protect Password:=Pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingColumns:=True, AllowFormattingRows:=True
    Next wSheet

End Sub
