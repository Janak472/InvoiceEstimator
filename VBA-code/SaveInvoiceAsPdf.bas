Attribute VB_Name = "Module1"
Sub SaveInvoiceAsPdf()
Dim NewFN As Variant
    NewFN = "C:\Users\nagwa\Desktop\Invoice Data\" & Range("B5") & ("_") & Range("G4").Value & ("_") & Range("G50").Value & ".pdf"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=NewFN, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
        
        
    



End Sub


