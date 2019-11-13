Sub CopyPaste_VLOOKUP_as_Values()

Set rng = Cells.Find(What:="Vlookup", _
After:=ActiveCell, LookIn:=xlFormulas, _
LookAt:=xlPart, SearchOrder:=xlByRows, _
SearchDirection:=xlNext, _
MatchCase:=False)

If Not rng Is Nothing Then
Do
rng.Formula = rng.Value

Set rng = Cells.FindNext(rng)
Loop Until rng Is Nothing

End If
End Sub
