Public Sub CopieRenome()
Dim Index As Integer
Index = 1
Dim nomsEtudiants As String
nomsEtudiants = "Liste déroulante"

Do
    Sheets("Base").Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = Sheets(nomsEtudiants).Range("A" & Index).Value
    Index = Index + 1
Loop While Sheets(nomsEtudiants).Range("A" & Index).Value <> ""

End Sub
