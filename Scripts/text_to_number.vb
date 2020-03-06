Sub text_to_number()
' This is a stupid macro to replace those stupid numbers in text format

Let x = 0

Do while x < 10:
    Selection.Replace What:=Cstr(x), Replacement:=Cstr(x), LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    x = x + 1
Loop
End Sub