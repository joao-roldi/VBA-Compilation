Function REPLACETEXTS(strInput As String, rngFind As Range, rngReplace As Range) As String

'Found on: https://www.howtoexcel.org/vba/how-to-find-and-replace-multiple-text-strings-within-a-text-string/

Dim strTemp As String
Dim strFind As String
Dim strReplace As String

Dim cellFind As Range

Dim lngColFind As Long
Dim lngRowFind As Long
Dim lngRowReplace As Long
Dim lngColReplace As Long

lngColFind = rngFind.Columns.Count
lngRowFind = rngFind.Rows.Count
lngColReplace = rngFind.Columns.Count
lngRowReplace = rngFind.Rows.Count

strTemp = strInput

If Not ((lngColFind = lngColReplace) And (lngRowFind = lngRowReplace)) Then
    REPLACETEXTS = CVErr(xlErrNA)
    Exit Function
End If

For Each cellFind In rngFind

    strFind = cellFind.Value
    strReplace = rngReplace(cellFind.Row - rngFind.Row + 1, cellFind.Column - rngFind.Column + 1).Value
    strTemp = Replace(strTemp, strFind, strReplace)

Next cellFind

REPLACETEXTS = strTemp

End Function