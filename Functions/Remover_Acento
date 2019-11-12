Function remove_accents(Caract As String)

'Function found in https://www.funcaoexcel.com.br/remover-acentos/

 Dim A As String
 Dim B As String
 Dim i As Integer
 Const AccChars = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
 Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
 For i = 1 To Len(AccChars)
 A = Mid(AccChars, i, 1)
 B = Mid(RegChars, i, 1)
 Caract = Replace(Caract, A, B)
 Next
 remove_accents = Caract
End Function
