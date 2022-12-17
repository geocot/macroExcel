Public Sub CopieRenome()
Dim Index As Integer
Index = 1

Do
    Sheets("Base").Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = Sheets("Liste déroulante").Range("A" & Index).Value
    Index = Index + 1
Loop While Sheets("Liste déroulante").Range("A" & Index).Value <> ""

End Sub
