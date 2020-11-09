VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelExportVersPDFs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ExportToPDFs()
' PDF Export Macro
' Changez C:\temp\ pour le répertoire de votre choix
' Changez les cellules à imprimer
' Sauvegarde chaque Worksheet vers un fichier PDF séparé.
' Changer l'aire d'impression sous Print Area
 
Dim ws As Worksheet
 
For Each ws In Worksheets
ws.Select
nm = ws.Name
With ActiveSheet.PageSetup
    .Zoom = False
    .PrintArea = "A1:I40"
    .Orientation = xlLandscape
    .PaperSize = xlPaperLegal
    .FitToPagesTall = 2
    .FitToPagesWide = 1
End With

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:="C:\temp\" + nm + ".pdf", _
Quality:=xlQualityStandard, IncludeDocProperties:=True, _
IgnorePrintAreas:=False, OpenAfterPublish:=False
 
Next ws
 
End Sub
