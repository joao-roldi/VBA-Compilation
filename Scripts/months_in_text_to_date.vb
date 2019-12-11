Sub monhts_in_text_to_date()
    Columns("C:C").Select
    Selection.Replace What:="jan", Replacement:="01/01/2019", LookAt:=xlPart , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="fev", Replacement:="02/01/2019", LookAt:=xlPart , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="mar", Replacement:="03/01/2019", LookAt:=xlPart , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="abr", Replacement:="04/01/2019", LookAt:=xlPart , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="mai", Replacement:="05/01/2019", LookAt:=xlPart , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="jun", Replacement:="06/01/2019", LookAt:=xlPart , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="jul", Replacement:="07/01/2019", LookAt:=xlPart , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="ago", Replacement:="08/01/2019", LookAt:=xlPart , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="set", Replacement:="09/01/2019", LookAt:=xlPart , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="out", Replacement:="10/01/2019", LookAt:=xlPart , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="nov", Replacement:="11/01/2019", LookAt:=xlPart , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:="dez", Replacement:="12/01/2019", LookAt:=xlPart , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub