Attribute VB_Name = "Module1"
   
Public Sub CopieRenome()
Dim Index As Integer

For Index = 45 To 58
    Sheets("Base").Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = Sheets("Liste déroulante").Range("A" & Index).Value
    
Next

End Sub
